using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using RestSharp;
using SAPbobsCOM;
using Pro4Soft.SapB1Integration.Infrastructure;
using Pro4Soft.SapB1Integration.Workers.Download;

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class PurchaseOrderUpload : IntegrationBase
    {
        public PurchaseOrderUpload(ScheduleSetting settings) : base(settings)
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
                    var pos = WebInvoke<List<PurchaseOrder>>("odata/PurchaseOrder", request =>
                    {
                        if(string.IsNullOrWhiteSpace(companySettings.ClientName))
                            request.AddQueryParameter("$filter", $@"IsWarehouseTransfer eq false and UploadDate eq null and PurchaseOrderState eq 'Closed' and ClientId eq null");
                        else
                            request.AddQueryParameter("$filter", $@"IsWarehouseTransfer eq false and UploadDate eq null and PurchaseOrderState eq 'Closed' and Client/Name eq '{companySettings.ClientName}'");
                        request.AddQueryParameter("$select", @"Id,PurchaseOrderNumber,ReferenceNumber");
                        request.AddQueryParameter("$expand", @"Lines($orderby=LineNumber;$select=Id,ReceivedQuantity,ReferenceNumber;$expand=Product($select=Id,Sku),LineDetails($orderby=LotNumber,SerialNumber)),Client($select=Id,Name)");
                        request.AddQueryParameter("$orderby", @"PurchaseOrderNumber");
                    }, Method.GET, null, "value");
                    if (!pos.Any())
                        continue;
                
                    foreach (var po in pos)
                    {
                        var b1PurchaseOrder = CurrentCompany.GetBusinessObject(BoObjectTypes.oPurchaseOrders) as Documents;
                        var b1Grpo = CurrentCompany.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes) as Documents;
                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
                        try
                        {
                            if (!b1PurchaseOrder.GetByKey(int.TryParse(po.ReferenceNumber, out var poKey) ? poKey : throw new Exception($"Cannot parse {po.ReferenceNumber}")))
                                throw new Exception($"Failed to retrieve PO DocEntry: {po.ReferenceNumber}");

                            b1Grpo.DocDate = DateTime.Now;
                            b1Grpo.CardCode = b1PurchaseOrder.CardCode;
                            var receiptLine = 0;
                            foreach (var line in po.Lines.Where(c => c.ReceivedQuantity > 0))
                            {
                                if (!item.GetByKey(line.Product.Sku))
                                    throw new Exception($"Failed to retrieve item: {line.Product.Sku}");

                                if (receiptLine != 0)
                                    b1Grpo.Lines.Add();
                                b1Grpo.Lines.SetCurrentLine(receiptLine++);

                                b1Grpo.Lines.BaseEntry = b1PurchaseOrder.DocEntry;
                                b1Grpo.Lines.BaseLine = int.TryParse(line.ReferenceNumber, out var refKey) ? refKey : throw new Exception($"Cannot parse {line.ReferenceNumber}");
                                b1Grpo.Lines.BaseType = (int) BoObjectTypes.oPurchaseOrders;
                                b1Grpo.Lines.Quantity = (double) line.ReceivedQuantity;

                                var uom = GetNumPerMsr(b1PurchaseOrder.DocEntry, b1Grpo.Lines.BaseLine, "POR1");
                                var qty = (double) line.ReceivedQuantity / uom;

                                b1Grpo.Lines.Quantity = qty;

                                var serialLine = 0;
                                if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                    foreach (var lineDetails in line.LineDetails)
                                    {
                                        if (serialLine != 0)
                                            b1Grpo.Lines.SerialNumbers.Add();
                                        b1Grpo.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                        b1Grpo.Lines.SerialNumbers.InternalSerialNumber = lineDetails.SerialNumber;
                                    }

                                if (item.ManageBatchNumbers != BoYesNoEnum.tYES)
                                    continue;

                                var batchLine = 0;
                                foreach (var adjustment in line.LineDetails.GroupBy(c => c.LotNumber))
                                {
                                    if (batchLine != 0)
                                        b1Grpo.Lines.BatchNumbers.Add();
                                    b1Grpo.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                    b1Grpo.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                    b1Grpo.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1Grpo.Lines.BatchNumbers.ExpiryDate;
                                    b1Grpo.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.ReceivedQuantity * (c.PacksizeEachCount ?? 1));
                                }
                            }

                            //Freight charges
                            var targetFreightLine = 0;
                            var baseFreightLine = 0;
                            while (baseFreightLine < b1PurchaseOrder.Expenses.Count)
                            {
                                b1PurchaseOrder.Expenses.SetCurrentLine(baseFreightLine++);

                                if (b1PurchaseOrder.Expenses.LineTotal == b1PurchaseOrder.Expenses.PaidToDate)
                                    continue;
                                if (targetFreightLine > 0)
                                    b1Grpo.Expenses.Add();
                                b1Grpo.Expenses.SetCurrentLine(targetFreightLine++);

                                b1Grpo.Expenses.BaseDocType = (int) b1PurchaseOrder.DocObjectCode;
                                b1Grpo.Expenses.BaseDocEntry = b1PurchaseOrder.DocEntry;
                                b1Grpo.Expenses.BaseDocLine = b1PurchaseOrder.Expenses.LineNum;
                            }

                            StagingTableLockUnlock(b1PurchaseOrder.DocEntry.ToString(), (int) BoObjectTypes.oPurchaseOrders, false);
                            if (b1Grpo.Add() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"PO upload [{po.PurchaseOrderNumber}] upload failed: {errorCode} - {errorMessage}");
                            }

                            WebInvoke("api/PurchaseOrderApi/CreateOrUpdate", null, Method.POST, new
                            {
                                po.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = true,
                                UploadMessage = (string) null,
                            });

                            b1PurchaseOrder.GetByKey(poKey);
                            if (b1PurchaseOrder.DocumentStatus == BoStatus.bost_Open)
                                StagingTableIncrementBackOrder(po.ReferenceNumber, (int) BoObjectTypes.oPurchaseOrders);
                            else
                                StagingTableDelete(po.ReferenceNumber, (int) BoObjectTypes.oPurchaseOrders);

                            Log($"PO: [{po.PurchaseOrderNumber}] for [{companySettings.ClientName ?? companySettings.CompanyDb}] uploaded");
                            runDownload = runDownload || b1PurchaseOrder.DocumentStatus == BoStatus.bost_Open;
                        }
                        catch (Exception e)
                        {
                            WebInvoke("api/PurchaseOrderApi/CreateOrUpdate", null, Method.POST, new
                            {
                                po.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = false,
                                UploadMessage = e.ToString()
                            });
                            LogError(e);
                            StagingTableLockUnlock(b1PurchaseOrder.DocEntry.ToString(), (int) BoObjectTypes.oPurchaseOrders, true);
                        }
                        finally
                        {
                            if (b1PurchaseOrder != null)
                                Marshal.ReleaseComObject(b1PurchaseOrder);
                            if (b1Grpo != null)
                                Marshal.ReleaseComObject(b1Grpo);
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
                    ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(PurchaseOrderDownload).FullName));
            }
        }
    }
}