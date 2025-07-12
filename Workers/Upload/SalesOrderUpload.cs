using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using RestSharp;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using Pro4Soft.SapB1Integration.Workers.Download;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class SalesOrderUpload : IntegrationBase
    {
        public SalesOrderUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            var runDownload = false;
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var sos = WebInvoke<List<SalesOrder>>("odata/PickTicket", request =>
                    {
                        if (string.IsNullOrWhiteSpace(companySettings.ClientName))
                            request.AddQueryParameter("$filter", $@"UploadDate eq null and PickTicketState eq 'Closed' and ClientId eq null");
                        else
                            request.AddQueryParameter("$filter", $@"UploadDate eq null and PickTicketState eq 'Closed' and Client/Name eq '{companySettings.ClientName}'");
                        request.AddQueryParameter("$select", @"Id,PickTicketNumber,ReferenceNumber");
                        request.AddQueryParameter("$orderby", @"PickTicketNumber");
                        request.AddQueryParameter("$expand", @"
Client($select=Name),
Totes($select=Id,Sscc18Code,CartonNumber;
      $expand=Lines($select=Id,PickedQuantity;
                    $expand=Product($select=Id,Sku),
                            PickTicketLine($select=Id,LineNumber,ReferenceNumber;$orderby=ReferenceNumber),
                            LineDetails($select=Id,PacksizeEachCount,LotNumber,SerialNumber,ExpiryDate,PickedQuantity)))");
                    }, Method.GET, null, "value");
                    if (!sos.Any())
                        continue;

                    foreach (var so in sos)
                    {
                        var b1SalesOrder = CurrentCompany.GetBusinessObject(BoObjectTypes.oOrders) as Documents;
                        var b1Delivery = CurrentCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes) as Documents;
                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;

                        try
                        {
                            if (!b1SalesOrder.GetByKey(int.TryParse(so.ReferenceNumber, out var soKey) ? soKey : throw new Exception($"Cannot parse {so.ReferenceNumber}")))
                                throw new Exception($"Failed to retrieve SO DocEntry: {so.ReferenceNumber}");

                            b1Delivery.DocDate = DateTime.Now;
                            b1Delivery.TaxDate = DateTime.Now;
                            b1Delivery.CardCode = b1SalesOrder.CardCode;

                            var deliveryLineNum = 0;
                            var lines = so.GetOrderLines();

                            for (var soLineNum = 0; soLineNum < b1SalesOrder.Lines.Count; soLineNum++)
                            {
                                b1SalesOrder.Lines.SetCurrentLine(soLineNum);

                                if (b1SalesOrder.Lines.TreeType == BoItemTreeTypes.iSalesTree)
                                {
                                    if (b1SalesOrder.Lines.RemainingOpenQuantity <= 0)
                                        continue;
                                    if (deliveryLineNum != 0)
                                        b1Delivery.Lines.Add();
                                    b1Delivery.Lines.SetCurrentLine(deliveryLineNum++);
                                    b1Delivery.Lines.BaseType = (int) BoObjectTypes.oOrders;
                                    b1Delivery.Lines.BaseLine = b1SalesOrder.Lines.LineNum;
                                    b1Delivery.Lines.BaseEntry = b1SalesOrder.DocEntry;
                                    b1Delivery.Lines.Quantity = b1SalesOrder.Lines.RemainingOpenQuantity;
                                    continue;
                                }

                                var line = lines.SingleOrDefault(c => b1SalesOrder.Lines.LineNum.ToString() == c.PickTicketLine.ReferenceNumber);
                                if (line == null)
                                    continue;

                                if (deliveryLineNum != 0)
                                    b1Delivery.Lines.Add();

                                b1Delivery.Lines.SetCurrentLine(deliveryLineNum++);
                                b1Delivery.Lines.BaseType = (int) BoObjectTypes.oOrders;
                                b1Delivery.Lines.BaseEntry = b1SalesOrder.DocEntry;
                                b1Delivery.Lines.BaseLine = b1SalesOrder.Lines.LineNum;
                                b1Delivery.Lines.Quantity = (double)line.PickedQuantity / GetNumPerMsr(b1Delivery.Lines.BaseEntry, b1Delivery.Lines.BaseLine, "RDR1");

                                if (!item.GetByKey(b1SalesOrder.Lines.ItemCode))
                                    continue;

                                if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                                {
                                    var batchLine = 0;
                                    foreach (var lineDetails in line.LineDetails.GroupBy(c => c.LotNumber))
                                    {
                                        if (batchLine != 0)
                                            b1Delivery.Lines.BatchNumbers.Add();
                                        b1Delivery.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                        b1Delivery.Lines.BatchNumbers.BatchNumber = lineDetails.Key;
                                        b1Delivery.Lines.BatchNumbers.Quantity = (double)lineDetails.Sum(c => c.PickedQuantity * (c.PacksizeEachCount ?? 1));
                                        //b1Delivery.Lines.BatchNumbers.ExpiryDate = lineDetails.FirstOrDefault().ExpiryDate?.DateTime ?? b1Delivery.Lines.BatchNumbers.ExpiryDate;
                                    }
                                }

                                if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                {
                                    var serialLine = 0;
                                    foreach (var detl in line.LineDetails.OrderBy(c=>c.SerialNumber))
                                    {
                                        if (serialLine != 0)
                                            b1Delivery.Lines.SerialNumbers.Add();
                                        b1Delivery.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                        b1Delivery.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(detl.SerialNumber, line.Product.Sku);
                                        b1Delivery.Lines.SerialNumbers.InternalSerialNumber = detl.SerialNumber;
                                        b1Delivery.Lines.SerialNumbers.Quantity = 1;
                                    }
                                }
                            }

                            //Freight charges
                            var targetFreightLine = 0;
                            var baseFreightLine = 0;
                            while (baseFreightLine < b1SalesOrder.Expenses.Count)
                            {
                                b1SalesOrder.Expenses.SetCurrentLine(baseFreightLine++);

                                if (Math.Abs(b1SalesOrder.Expenses.LineTotal - b1SalesOrder.Expenses.PaidToDate) < 0.00001)
                                    continue;
                                if (targetFreightLine > 0)
                                    b1Delivery.Expenses.Add();
                                b1Delivery.Expenses.SetCurrentLine(targetFreightLine++);

                                b1Delivery.Expenses.BaseDocType = (int) b1SalesOrder.DocObjectCode;
                                b1Delivery.Expenses.BaseDocEntry = b1SalesOrder.DocEntry;
                                b1Delivery.Expenses.BaseDocLine = b1SalesOrder.Expenses.LineNum;
                            }

                            StagingTableLockUnlock(b1SalesOrder.DocEntry.ToString(), (int)BoObjectTypes.oOrders, false);
                            if (b1Delivery.Add() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                Log(Utils.SerializeToStringJson(so, Formatting.Indented));
                                throw new Exception($"SO [{so.PickTicketNumber}] upload failed: {errorCode} - {errorMessage}");
                            }

                            WebInvoke("api/PickTicketApi/CreateOrUpdate", null, Method.POST, new
                            {
                                so.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = true,
                                UploadMessage = (string) null,
                            });

                            b1SalesOrder.GetByKey(soKey);
                            if (b1SalesOrder.DocumentStatus == BoStatus.bost_Open)
                                StagingTableIncrementBackOrder(so.ReferenceNumber, (int) BoObjectTypes.oOrders);
                            else
                                StagingTableDelete(so.ReferenceNumber, (int) BoObjectTypes.oOrders);

                            Log($"SO: [{so.PickTicketNumber}] for [{companySettings.ClientName ?? companySettings.CompanyDb}] uploaded");
                            runDownload = runDownload || b1SalesOrder.DocumentStatus == BoStatus.bost_Open;
                        }
                        catch (Exception e)
                        {
                            WebInvoke("api/PickTicketApi/CreateOrUpdate", null, Method.POST, new
                            {
                                so.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = false,
                                UploadMessage = e.ToString()
                            });
                            LogError(e);
                            StagingTableLockUnlock(b1SalesOrder.DocEntry.ToString(), (int)BoObjectTypes.oOrders, true);
                        }
                        finally
                        {
                            if (b1SalesOrder != null)
                                Marshal.ReleaseComObject(b1SalesOrder);
                            if (b1Delivery != null)
                                Marshal.ReleaseComObject(b1Delivery);
                            if (item != null)
                                Marshal.ReleaseComObject(item);
                        }
                    }
                }
                catch (Exception e)
                {
                    LogError(e);
                }
                finally
                {
                    Disconnect();
                }

                if(runDownload)
                    ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(SalesOrderDownload).FullName));
            }
        }
    }
}