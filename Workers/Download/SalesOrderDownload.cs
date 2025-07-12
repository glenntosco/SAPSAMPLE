using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

// ReSharper disable AccessToModifiedClosure
// ReSharper disable CompareOfFloatsByEqualityOperator

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class SalesOrderDownload : IntegrationBase
    {
        public SalesOrderDownload(ScheduleSetting settings) : base(settings)
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
    ORDR.DocEntry,
    ORDR.DocNum,
    ORDR.CardCode,
    ORDR.CancelDate,
	ORDR.ReqDate,
    OCRD.CntctPrsn,
    OCRD.E_Mail,
	OCRD.Phone1,
    RDR1.LineNum,
    RDR1.ItemCode,
    RDR1.OpenQty,
    RDR1.WhsCode,
    RDR1.NumPerMsr,
    RDR1.TreeType,
    RDR12.StreetS,
    RDR12.BlockS,
    RDR12.BuildingS,
    RDR12.CityS,
    RDR12.ZipCodeS,
    RDR12.StateS,
    RDR12.CountryS,
    ORDR.DocStatus,
    OITM.InvntItem,
    ORDR.CardName,
    ORDR.ShipToCode,
    RDR12.StreetNoS,
    ORDR.Comments,
    [@P4S_EXPORT].U_BackOrder
from 
    ORDR
    join OCRD on OCRD.CardCode = ORDR.CardCode
    left outer join [@P4S_Export] on CAST(ORDR.DocEntry AS NVARCHAR(10)) = [@P4S_Export].U_DocId and ORDR.ObjType = [@P4S_Export].U_ObjType
    left outer join RDR1 on ORDR.DocEntry = RDR1.DocEntry
    left outer join RDR12 on ORDR.DocEntry = RDR12.DocEntry and ORDR.ObjType = RDR12.ObjectType
    left outer join OITM on OITM.ItemCode = RDR1.ItemCode
where 
    [@P4S_Export].U_ExportFlag = 0
    and ORDR.Confirmed = 'Y'
    and RDR1.LineNum IS NOT NULL
    and OITM.InvntItem = 'Y'
    and U_ObjType = {(int)BoObjectTypes.oOrders}");

                    var clientId = GetClientId(companySettings.ClientName);

                    foreach (var customerGroup in data.GroupBy(c => c.CardCode))
                    {
                        var customer = customerGroup.First();

                        var customerId = IdLookup("odata/Customer", request =>
                        {
                            request.AddQueryParameter("$filter", $"CustomerCode eq '{customer.CardCode}' and ClientId eq " + (clientId?.ToString() ?? "null"));
                        });
                        if (customerId == null)
                        {
                            customerId = Guid.TryParse(WebInvoke<dynamic>("api/CustomerApi/CreateOrUpdate", null, Method.POST, new
                            {
                                ClientId = clientId,
                                CustomerCode = customer.CardCode,
                                CompanyName = customer.CardName,
                                ContactPerson = customer.CntctPrsn,
                                Email = customer.E_Mail,
                                Phone = customer.Phone1
                            }).Id.ToString(), out Guid vendId) ? vendId : throw new BusinessWebException($"Customer [{customer.CardCode}] could not be created");
                            Log($"Customer [{customer.CardCode}] created");
                        }

                        foreach (var orderGroup in customerGroup.GroupBy(c => c.DocEntry))
                        {
                            var order = orderGroup.First();
                            try
                            {
                                var orderCode = $"{order.DocNum}" + (order.U_BackOrder > 0 ? $"-{order.U_BackOrder}" : string.Empty);
                                if (order.DocStatus == "L" || order.DocStatus == "C")
                                {
                                    var count = WebInvoke<dynamic>("api/PickTicketApi/DeleteByCode", request =>
                                    {
                                        request.AddQueryParameter("code", orderCode);
                                        request.AddQueryParameter("clientId", clientId?.ToString());
                                    });
                                    StagingTableDelete(order.DocEntry.ToString(), (int)BoObjectTypes.oOrders);
                                    if(count > 0)
                                        Log($"SO: [{orderCode}] deleted!");
                                    continue;
                                }

                                var payload = new
                                {
                                    ClientId = clientId,
                                    CustomerId = customerId,
                                    WarehouseCode = order.WhsCode,
                                    order.CancelDate,
                                    RequiredDate = order.ReqDate,
                                    PickTicketNumber = orderCode,
                                    ReferenceNumber = order.DocEntry,
                                    ShipToName = order.ShipToCode,
                                    ShipToAddress1 = string.Join(" ", new List<string> {order.BuildingS, order.StreetNoS, order.StreetS}.Where(c => !string.IsNullOrWhiteSpace(c))),
                                    ShipToAddress2 = order.BlockS,
                                    ShipToCity = order.CityS,
                                    ShipToStateProvince = order.StateS,
                                    ShipToZipPostal = order.ZipCodeS,
                                    ShipToCountry = order.CountryS,
                                    order.Comments,
                                    Lines = orderGroup.GroupBy(c => c.LineNum)
                                        .Select(c => c.First())
                                        .Where(c => c.OpenQty > 0)
                                        .Where(c => c.TreeType == "N" || order.U_BackOrder == 0)
                                        .Select(c => new
                                        {
                                            Sku = c.ItemCode,
                                            ReferenceNumber = c.LineNum,
                                            Packsize = c.NumPerMsr > 1 ? c.NumPerMsr : null,
                                            NumberOfPacks = c.NumPerMsr > 1 ? (int?) c.OpenQty : null,
                                            OrderedQuantity = c.OpenQty * c.NumPerMsr
                                        }).ToList()
                                };

                                if (payload.Lines.Count > 0)
                                {
                                    WebInvoke<dynamic>("api/PickTicketApi/CreateOrUpdate", null, Method.POST, payload);
                                    Log($"SO: [{orderCode}] downloaded!");
                                }
                                else
                                    Log($"SO: [{orderCode}] skipped, no valid lines on an order!");
                                StagingTableMarkDownloaded(order.DocEntry.ToString(), (int)BoObjectTypes.oOrders);
                            }
                            catch (Exception ex)
                            {
                                LogError($@"SO: [{order.DocNum}] download failed: {ex}
----------
{Utils.SerializeToStringJson(order, Formatting.Indented)}");
                                StagingTableMarkDownloaded(order.DocEntry.ToString(), (int)BoObjectTypes.oOrders, ex.ToString().Truncate(254));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError($"SO failed to download: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }
}