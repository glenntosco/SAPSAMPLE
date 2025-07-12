using System;
using System.Linq;
using Newtonsoft.Json;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

// ReSharper disable AccessToModifiedClosure

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class PurchaseOrderDownload : IntegrationBase
    {
        public PurchaseOrderDownload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var data = ExecuteReader($@"
select 
    OPOR.DocEntry,
    OPOR.DocNum,
    OPOR.CardCode,
    OPOR.CancelDate,
	OPOR.ReqDate,
    OPOR.CardName,
    OCRD.CntctPrsn,
    OCRD.E_Mail,
	OCRD.Phone1,
    POR1.LineNum, 
    POR1.ItemCode, 
    POR1.NumPerMsr, 
    POR1.OpenQty, 
    [@P4S_Export].U_BackOrder,
    OPOR.DocStatus, 
    POR1.WhsCode, 
    OPOR.Comments
from 
    [@P4S_Export]
	join OPOR on CAST(OPOR.DocEntry AS NVARCHAR(10)) = [@P4S_Export].U_DocId and opor.ObjType = [@P4S_Export].U_ObjType 
    join POR1 on OPOR.DocEntry = POR1.DocEntry
	join OCRD on OCRD.CardCode = opor.CardCode
    join OITM on OITM.ItemCode = POR1.ItemCode
where 
    U_ExportFlag = 0    
    and OPOR.Confirmed = 'Y'
    and POR1.LineNum IS NOT NULL
    and OITM.InvntItem = 'Y'
    and U_ObjType = {(int)BoObjectTypes.oPurchaseOrders}");

                    var clientId = GetClientId(companySettings.ClientName);

                    foreach (var vendorGroup in data.GroupBy(c => c.CardCode))
                    {
                        var vendor = vendorGroup.First();

                        var vendorId = IdLookup("odata/Vendor", request =>
                        {
                            request.AddQueryParameter("$filter", $"VendorCode eq '{vendor.CardCode}' and ClientId eq " + (clientId?.ToString() ?? "null"));
                        });
                        if (vendorId == null)
                        {
                            vendorId = Guid.TryParse(WebInvoke<dynamic>("api/VendorApi/CreateOrUpdate", null, Method.POST, new
                            {
                                ClientId = clientId,
                                VendorCode = vendor.CardCode,
                                CompanyName = vendor.CardName,
                                ContactPerson = vendor.CntctPrsn,
                                Email = vendor.E_Mail,
                                Phone = vendor.Phone1
                            }).Id.ToString(), out Guid vendId) ? vendId : throw new BusinessWebException($"Vendor [{vendor.CardCode}] could not be created");
                            Log($"Vendor [{vendor.CardCode}] created");
                        }

                        foreach (var orderGroup in vendorGroup.GroupBy(c => c.DocEntry))
                        {
                            var order = orderGroup.First();
                            try
                            {
                                var orderCode = $"{order.DocNum}" + (order.U_BackOrder > 0 ? $"-{order.U_BackOrder}" : string.Empty);
                                if (order.DocStatus == "L" || order.DocStatus == "C")
                                {
                                    var count = WebInvoke<dynamic>("api/PurchaseOrderApi/DeleteByCode", request =>
                                    {
                                        request.AddQueryParameter("code", orderCode);
                                        request.AddQueryParameter("clientId", clientId?.ToString());
                                    });
                                    StagingTableDelete(order.DocEntry.ToString(), (int) BoObjectTypes.oPurchaseOrders);
                                    if(count > 0)
                                        Log($"PO: [{orderCode}] deleted!");
                                    continue;
                                }

                                var payload = new
                                {
                                    ClientId = clientId,
                                    VendorId = vendorId,
                                    WarehouseCode = order.WhsCode,
                                    order.CancelDate,
                                    RequiredDate = order.ReqDate,
                                    PurchaseOrderNumber = orderCode,
                                    ReferenceNumber = order.DocEntry,
                                    order.Comments,
                                    Lines = orderGroup.GroupBy(c => c.LineNum)
                                        .Select(c => c.First())
                                        .Where(c => c.OpenQty > 0)
                                        .Select(c => new
                                        {
                                            Sku = c.ItemCode,
                                            ReferenceNumber = c.LineNum,
                                            Packsize = c.NumPerMsr,
                                            NumberOfPacks = (int) c.OpenQty,
                                            OrderedQuantity = c.OpenQty * c.NumPerMsr
                                        }).ToList()
                                };
                                if (payload.Lines.Count > 0)
                                {
                                    WebInvoke<dynamic>("api/PurchaseOrderApi/CreateOrUpdate", null, Method.POST, payload);
                                    Log($"PO: [{orderCode}] downloaded!");
                                }
                                else
                                    Log($"PO: [{orderCode}] skipped, no valid lines on an order!");

                                StagingTableMarkDownloaded(order.DocEntry.ToString(), (int) BoObjectTypes.oPurchaseOrders);
                            }
                            catch (Exception ex)
                            {
                                LogError($@"PO: [{order.DocNum}] download failed: {ex}
----------
{Utils.SerializeToStringJson(order, Formatting.Indented)}");
                                StagingTableMarkDownloaded(order.DocEntry.ToString(), (int) BoObjectTypes.oPurchaseOrders, ex.ToString().Truncate(254));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError($"PO failed to download: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }
}