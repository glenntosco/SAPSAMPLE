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
    public class WorkOrderDownload : IntegrationBase
    {
        public WorkOrderDownload(ScheduleSetting settings) : base(settings)
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
    OWOR.Warehouse,	
    OWOR.ItemCode as BuildSku,
    OWOR.DocNum,	
    OWOR.DocEntry,	
	OWOR.PlannedQty - OWOR.CmpltQty as QuantityToBuild,
    OWOR.Comments,
    OWOR.Status,
    WOR1.LineNum,
    U_BackOrder,
    WOR1.ItemCode as CompSku,
	WOR1.PlannedQty - WOR1.IssuedQty as QuantityToConsume
from OWOR
	left outer join WOR1 on WOR1.DocEntry = OWOR.DocEntry
	left outer join [@P4S_EXPORT] on CAST(OWOR.DocEntry AS NVARCHAR(10)) = [@P4S_EXPORT].U_DocId
where
	OWOR.Status in ('R','C','L') and
	OWOR.Type = 'S' and
	WOR1.IssueType = 'M' and
    [@P4S_Export].U_ExportFlag = 0 and
	[@P4S_EXPORT].U_ObjType = {(int)BoObjectTypes.oProductionOrders}");
                    var clientId = GetClientId(companySettings.ClientName);

                    foreach (var orderGroup in data.GroupBy(c => c.DocEntry))
                    {
                        var order = orderGroup.First();
                        var orderCode = $"{order.DocNum}" + (order.U_BackOrder > 0 ? $"-{order.U_BackOrder}" : string.Empty);
                        try
                        {
                            if (order.Status == "C" || order.Status == "L")
                            {
                                var count = WebInvoke<dynamic>("api/WorkOrderApi/DeleteByCode", request =>
                                {
                                    request.AddQueryParameter("code", orderCode);
                                    request.AddQueryParameter("clientId", clientId?.ToString());
                                });
                                StagingTableDelete(order.DocEntry.ToString(), (int)BoObjectTypes.oProductionOrders);
                                if (count > 0)
                                    Log($"WO: [{orderCode}] deleted!");
                                continue;
                            }

                            var payload = new
                            {
                                ClientId = clientId,
                                WarehouseCode = order.Warehouse,
                                Sku = order.BuildSku,
                                WorkOrderNumber = orderCode,
                                ReferenceNumber = order.DocEntry,
                                Quantity = order.QuantityToBuild,
                                order.Comments,
                                Lines = orderGroup.GroupBy(c => c.LineNum)
                                    .Select(c => c.First())
                                    .Where(c => c.QuantityToConsume > 0)
                                    .Select(c => new
                                    {
                                        Sku = c.CompSku,
                                        ReferenceNumber = c.LineNum,
                                        Quantity = c.QuantityToConsume
                                    }).ToList()
                            };

                            if (payload.Lines.Count > 0 && payload.Quantity > 0)
                            {
                                WebInvoke<dynamic>("api/WorkOrderApi/CreateOrUpdate", null, Method.POST, payload);
                                Log($"WO: [{orderCode}] downloaded!");
                            }
                            else
                                Log($"WO: [{orderCode}] skipped. Invalid Production Quantity or no valid lines on an order!");

                            StagingTableMarkDownloaded(order.DocEntry.ToString(), (int)BoObjectTypes.oProductionOrders);
                        }
                        catch (Exception ex)
                        {
                            LogError($@"WO: [{order.DocNum}] download failed: {ex}
----------
{Utils.SerializeToStringJson(order, Formatting.Indented)}");
                            StagingTableMarkDownloaded(order.DocEntry.ToString(), (int)BoObjectTypes.oProductionOrders, ex.ToString().Truncate(254));
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError($"WO failed to download: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }
}