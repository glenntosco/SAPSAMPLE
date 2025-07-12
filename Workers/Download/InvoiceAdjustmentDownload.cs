using System;
using System.Collections.Generic;
using System.Linq;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

// ReSharper disable AccessToModifiedClosure
// ReSharper disable CompareOfFloatsByEqualityOperator

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class InvoiceAdjustmentDownload : IntegrationBase
    {
        public InvoiceAdjustmentDownload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(companySettings.InvoiceAdjustmentDownloadBin))
                        continue;

                    SetCompany(companySettings);
                    var data = ExecuteReader($@"
select 
	OINV.DocNum,
    OINV.DocEntry,
	INV1.ItemCode,
	INV1.OpenQty as Quantity,
	INV1.NumPerMsr as Packsize,
	INV1.WhsCode,
	ITL1.Quantity * -1 as DetaillQuantity,
	OBTN.DistNumber as LotNumber,
	OSRN.DistNumber as SerialNumber
from OINV
	left outer join INV1 on INV1.DocEntry = OINV.DocEntry
	left outer join OITL on INV1.DocEntry = OITL.DocEntry and INV1.LineNum = OITL.DocLine and OITL.DocType = '{(int)BoObjectTypes.oInvoices}'
	left outer join ITL1 on OITL.LogEntry = ITL1.LogEntry
	left outer join OSRN on ITL1.SysNumber = OSRN.SysNumber and OITL.ItemCode = OSRN.ItemCode
	left outer join OBTN on ITL1.SysNumber = OBTN.SysNumber and OITL.ItemCode = OBTN.ItemCode
	left outer join [@P4S_EXPORT] on CAST(OINV.DocEntry AS NVARCHAR(10)) = [@P4S_EXPORT].U_DocId
where 
	INV1.BaseEntry is null and 
	[@P4S_EXPORT].U_ExportFlag = 0 and 
	[@P4S_EXPORT].U_BackOrder = 0");
                    if (!data.Any())
                        continue;

                    var binDetls = GetBinDetails(companySettings.InvoiceAdjustmentDownloadBin);
                    var prodCache = new Dictionary<string, FullProduct>();
                    foreach (var docGroup in data.GroupBy(c => c.DocNum))
                    {
                        string docNum = docGroup.Key.ToString();
                        string docEntry = docGroup.First().DocEntry.ToString();
                        try
                        {
                            var payload = new List<ProductWarehouseOp>();
                            foreach (var item in docGroup)
                            {
                                string itemCode = item.ItemCode.ToString();
                                if (!prodCache.TryGetValue(itemCode, out var existingProd))
                                {
                                    existingProd = WebInvoke<List<FullProduct>>("odata/Product", request =>
                                    {
                                        request.AddQueryParameter("$filter", string.IsNullOrWhiteSpace(companySettings.ClientName) ? 
                                                $"Sku eq '{item.ItemCode}' and ClientId eq null" : 
                                                $"Sku eq '{item.ItemCode}' and Client/Name eq '{companySettings.ClientName}'");
                                        request.AddQueryParameter("$expand", "Packsizes");
                                    }, Method.GET, null, "value").SingleOrDefault();
                                    prodCache[item.ItemCode] = existingProd ?? throw new Exception($"SKU: {item.ItemCode} is not setup in P4W");
                                }

                                var newLine = new ProductWarehouseOp
                                {
                                    ProductId = existingProd.Id,
                                    BinId = binDetls.Id,
                                    ReferenceType = "Invoice",
                                    ReferenceCode = item.DocNum.ToString(),
                                    Quantity = item.Quantity
                                };

                                if (existingProd.IsExpiryControlled)
                                    throw new NotSupportedException($"Expiry controlled items are not supported on Invoice download");

                                if (existingProd.IsPacksizeControlled)
                                {
                                    var packsize = existingProd.Packsizes.SingleOrDefault(c => c.EachCount == item.Packsize);
                                    if (packsize == null)
                                        throw new Exception($"Packsize with [{item.Packsize}] for SKU [{item.ItemCode}] is not setup");
                                    newLine.PacksizeId = packsize.Id;
                                }

                                if (existingProd.IsLotControlled)
                                {
                                    newLine.LotNumber = item.LotNumber;
                                    newLine.Quantity = item.DetaillQuantity;
                                }
                                else if (existingProd.IsSerialControlled)
                                {
                                    newLine.SerialNumber = item.SerialNumber;
                                    newLine.Quantity = 1;
                                }

                                payload.Add(newLine);
                            }
                            if (payload.Any())
                            {
                                var resp = Singleton<EntryPoint>.Instance.WebInvoke("api/WarehouseOperations/AdjustOutBatch", null, Method.POST, payload);
                                if (!resp.IsSuccessful)
                                    throw new BusinessWebException(resp.StatusCode, resp.Content);
                                Log($"Invoice adjustment: [{docNum}] downloaded!");
                            }
                            
                            StagingTableMarkDownloaded(docEntry, (int)BoObjectTypes.oInvoices);
                            StagingTableIncrementBackOrder(docEntry, (int)BoObjectTypes.oInvoices);
                        }
                        catch (Exception ex)
                        {
                            LogError($@"Invoice adjustment: [{docNum}] download failed: {ex}");
                            StagingTableMarkDownloaded(docEntry, (int) BoObjectTypes.oInvoices, ex.ToString().Truncate(254));
                            StagingTableIncrementBackOrder(docEntry, (int)BoObjectTypes.oInvoices);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Orders failed to load: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }

    public class ProductWarehouseOp
    {
        public Guid? BinId { get; set; }
        public Guid ProductId { get; set; }
        public decimal Quantity { get; set; }
        public string Reason { get; set; }

        public Guid? PacksizeId { get; set; }
        public string LotNumber { get; set; }
        public string SerialNumber { get; set; }

        public string ReferenceCode { get; set; }
        public Guid? ReferenceId { get; set; }
        public string ReferenceType { get; set; }

        public ProductWarehouseOp Clone()
        {
            return new ProductWarehouseOp
            {
                BinId = BinId,
                ProductId = ProductId,
                Quantity = Quantity,
                Reason = Reason,
                PacksizeId = PacksizeId,
                LotNumber = LotNumber,
                SerialNumber = SerialNumber,
                ReferenceCode = ReferenceCode,
                ReferenceType = ReferenceType,
                ReferenceId = ReferenceId,
            };
        }
    }
}