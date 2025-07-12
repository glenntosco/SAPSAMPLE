using System;
using System.Linq;
using Newtonsoft.Json;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

// ReSharper disable AccessToModifiedClosure
// ReSharper disable CompareOfFloatsByEqualityOperator

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class WarehouseTransferDownload : IntegrationBase
    {
        public WarehouseTransferDownload(ScheduleSetting settings) : base(settings)
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
    OWTQ.DocNum,
    OWTQ.DocStatus,
    OWTQ.DocEntry,
    OWTQ.ObjType,
    OWTQ.Filler,
    OWTQ.ToWhsCode,
    OWTQ.JrnlMemo,
    WTQ1.LineNum,
    WTQ1.OpenQty,
    WTQ1.ItemCode,
    WTQ1.NumPerMsr,
    WTQ1.FromWhsCod,
    WTQ1.WhsCode
from 
    OWTQ
left outer join WTQ1 on OWTQ.DocEntry = WTQ1.DocEntry
left outer join [@P4S_EXPORT] on CAST(OWTQ.DocEntry AS NVARCHAR(10)) = [@P4S_EXPORT].U_DocId and OWTQ.ObjType = [@P4S_EXPORT].U_ObjType
where 
    [@P4S_EXPORT].U_ExportFlag = 0  and WTQ1.LineNum IS NOT NULL");

                    var clientId = GetClientId(companySettings.ClientName);
                    foreach (var orderGroup in data.GroupBy(c => c.DocEntry))
                    {
                        var order = orderGroup.First();
                        try
                        {
                            if (order.DocStatus == "L" || order.DocStatus == "C")
                            {
                                var count = WebInvoke<dynamic>("api/PickTicketApi/DeleteByCode", request =>
                                {
                                    request.AddQueryParameter("code", order.DocNum.ToString());
                                    request.AddQueryParameter("clientId", clientId?.ToString());
                                });
                                StagingTableDelete(order.DocEntry.ToString(), (int) BoObjectTypes.oInventoryTransferRequest);
                                if (count > 0)
                                    Log($"WH Transfer: [{order.DocNum}] deleted!");
                                continue;
                            }

                            var payload = new
                            {
                                ClientId = clientId,
                                WarehouseCode = order.Filler,
                                IsWarehouseTransfer = true,
                                ToWarehouseCode = order.ToWhsCode,
                                PickTicketNumber = order.DocNum,
                                ReferenceNumber = order.DocEntry,
                                order.JrnlMemo,
                                Lines = orderGroup.GroupBy(c => c.LineNum)
                                    .Select(c => c.First())
                                    .Where(c => c.OpenQty > 0)
                                    .Select(c => new
                                    {
                                        Sku = c.ItemCode,
                                        ReferenceNumber = c.LineNum,
                                        //Packsize = c.NumPerMsr,
                                        //NumberOfPacks = (int) c.OpenQty,
                                        OrderedQuantity = c.OpenQty * c.NumPerMsr
                                    }).ToList()
                            };

                            if (payload.Lines.Count > 0)
                            {
                                WebInvoke<dynamic>("api/PickTicketApi/CreateOrUpdate", null, Method.POST, payload);
                                Log($"WH Transfer: [{order.DocNum}] downloaded!");
                            }
                            else
                                Log($"WH Transfer: [{order.DocNum}] skipped, no valid lines on an order!");

                            StagingTableMarkDownloaded(order.DocEntry.ToString(), (int) BoObjectTypes.oInventoryTransferRequest);
                        }
                        catch (Exception ex)
                        {
                            LogError($@"WH Transfer: [{order.DocNum}] download failed: {ex}
----------
{Utils.SerializeToStringJson(order, Formatting.Indented)}");
                            StagingTableMarkDownloaded(order.DocEntry.ToString(), (int) BoObjectTypes.oInventoryTransferRequest, ex.ToString().Truncate(254));
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError($"WH Transfers failed to download: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }
}