using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using SAPbobsCOM;


namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class InventorySyncUpload : IntegrationBase
    {
        public InventorySyncUpload(ScheduleSetting settings) : base(settings)
        {

        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                SetCompany(companySettings);

                var service = CurrentCompany.GetCompanyService();
                var countService = service.GetBusinessService(ServiceTypes.InventoryCountingsService) as IInventoryCountingsService;
                var b1CountingDocument = countService.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting) as InventoryCounting;
                var b1Item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;

                try
                {
                    var allItems = new List<InventoryHelper>();
                    var page = 0;
                    var time = new Stopwatch();
                    Log($"Retrieving counts from P4W for company [{companySettings.CompanyDb}]");
                    var count = 0;
                    const int pageSize = 50;
                    time.Start();
                    while (true)
                    {
                        var serverVal = WebInvoke<List<InventoryHelper>>("api/Report/GetInventory", request =>
                        {
                            request.AddQueryParameter("client", companySettings.ClientName);
                            request.AddQueryParameter("page", page++.ToString());
                            request.AddQueryParameter("pageSize", pageSize.ToString());
                        });
                        count += serverVal.Count;
                        Log($"{count} items read. Elapsed time {time.Elapsed}");
                        allItems.AddRange(serverVal);
                        if (serverVal.Count < pageSize)
                            break;
                    }
                    Log($"A total of [{allItems.Count}] items read. Elapsed time {time.Elapsed}");

                    var packsizeMap = new Dictionary<string, string>();

                    var progressCount = 0;
                    Log("Building SAP Document...");
                    time.Restart();
                    foreach (var whGroup in allItems.GroupBy(c => c.Warehouse))
                    {
                        foreach (var skuGroup in whGroup.GroupBy(c => c.Sku).OrderBy(c => c.Key))
                        {
                            progressCount++;
                            b1Item.GetByKey(skuGroup.Key);
                            var item = skuGroup.First();
                            var line = b1CountingDocument.InventoryCountingLines.Add();
                            if (item.IsPacksizeControlled)
                            {
                                if (!packsizeMap.TryGetValue(item.Sku, out var packsize))
                                {
                                    packsize = ExecuteScalar<string>($@"
select 
	UomCode
from 
	ouom
	join ugp1 on ouom.UomEntry = ugp1.UomEntry
	join ougp on ougp.UgpEntry = ugp1.UgpEntry
	join oitm on oitm.UgpEntry = ougp.UgpEntry
where
	oitm.ItemCode = '{item.Sku}'
	and BaseQty/AltQty = 1");
                                    packsizeMap[item.Sku] = packsize;
                                }

                                if (string.IsNullOrWhiteSpace(packsize))
                                    throw new Exception($"SKU [{item.Sku}] is missing an Each packsize");
                                line.UoMCode = packsize;
                            }

                            line.ItemCode = b1Item.ItemCode;
                            line.WarehouseCode = whGroup.Key;
                            line.CountedQuantity = (double)skuGroup.Sum(c => c.Quantity);
                            line.Counted = BoYesNoEnum.tYES;

                            var details = skuGroup.SelectMany(c => c.Details).ToList();

                            if (b1Item.ManageBatchNumbers == BoYesNoEnum.tYES)
                            {
                                foreach (var detail in details)
                                {
                                    var batchLine = line.InventoryCountingBatchNumbers.Add();
                                    batchLine.BatchNumber = detail.LotNumber;
                                    //if(detail.Expiry != null)
                                    //    batchLine.ExpiryDate = detail.Expiry.Value.Date;
                                    batchLine.Quantity = (double)detail.Quantity;
                                }
                            }

                            if (b1Item.ManageSerialNumbers == BoYesNoEnum.tYES)
                            {
                                foreach (var detail in details)
                                {
                                    var serialLine = line.InventoryCountingSerialNumbers.Add();
                                    serialLine.Quantity = 1;
                                    serialLine.InternalSerialNumber = detail.SerialNumber;
                                    serialLine.ManufacturerSerialNumber = detail.SerialNumber;
                                }
                            }

                            if (progressCount % 50 == 0)
                                Log($"Added [{progressCount}/{allItems.Count}] lines to a document. Elapsed time {time.Elapsed}");
                        }
                    }
                    Log($"A total of [{b1CountingDocument.InventoryCountingLines.Count}] lines created on a document. Elapsed time {time.Elapsed}");
                    if (b1CountingDocument.InventoryCountingLines.Count == 0)
                        return;
                    Log($"Posting a document...");
                    time.Restart();
                    countService.Add(b1CountingDocument);
                    Log($"Document posted. Elapsed time {time.Elapsed}");
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
                finally
                {
                    if (service != null)
                        Marshal.ReleaseComObject(service);
                    if (countService != null)
                        Marshal.ReleaseComObject(countService);
                    if (b1CountingDocument != null)
                        Marshal.ReleaseComObject(b1CountingDocument);
                    if (b1Item != null)
                        Marshal.ReleaseComObject(b1Item);
                    Disconnect();
                }
            }
        }
    }
}