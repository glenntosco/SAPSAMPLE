using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using RestSharp;
using SAPbobsCOM;
using Pro4Soft.SapB1Integration.Infrastructure;

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class WarehouseTransferUpload : IntegrationBase
    {
        public WarehouseTransferUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var pos = WebInvoke<List<PurchaseOrder>>("odata/PurchaseOrder", request =>
                    {
                        if (string.IsNullOrWhiteSpace(companySettings.ClientName))
                            request.AddQueryParameter("$filter", $@"IsWarehouseTransfer eq true and UploadDate eq null and PurchaseOrderState eq 'Closed' and ClientId eq null");
                        else
                            request.AddQueryParameter("$filter", $@"IsWarehouseTransfer eq true and UploadDate eq null and PurchaseOrderState eq 'Closed' and Client/Name eq '{companySettings.ClientName}'");
                        request.AddQueryParameter("$select", @"Id,PurchaseOrderNumber,ReferenceNumber");
                        request.AddQueryParameter("$expand", @"Lines($orderby=LineNumber;$select=Id,ReceivedQuantity,ReferenceNumber;$expand=Product($select=Id,Sku),LineDetails($orderby=LotNumber,SerialNumber)),Client($select=Id,Name)");
                        request.AddQueryParameter("$orderby", @"PurchaseOrderNumber");
                    }, Method.GET, null, "value");
                    if (!pos.Any())
                        continue;
                
                    foreach (var po in pos)
                    {
                        var b1TransferRequest = CurrentCompany.GetBusinessObject(BoObjectTypes.oInventoryTransferRequest) as StockTransfer;
                        var b1Transfer = CurrentCompany.GetBusinessObject(BoObjectTypes.oStockTransfer) as StockTransfer;
                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
                        try
                        {
                            if (!b1TransferRequest.GetByKey(int.TryParse(po.ReferenceNumber, out var poKey) ? poKey : throw new Exception($"Cannot parse {po.ReferenceNumber}")))
                                throw new Exception($"Failed to retrieve Transfer Request DocEntry: {po.ReferenceNumber}");

                            b1Transfer.DocDate = DateTime.Now;
                            b1Transfer.TaxDate = DateTime.Now;
                            
                            var receiptLine = 0;
                            foreach (var line in po.Lines.Where(c => c.ReceivedQuantity > 0))
                            {
                                if (!item.GetByKey(line.Product.Sku))
                                    throw new Exception($"Failed to retrieve item: {line.Product.Sku}");

                                if (receiptLine != 0)
                                    b1Transfer.Lines.Add();
                                b1Transfer.Lines.SetCurrentLine(receiptLine++);

                                b1Transfer.Lines.BaseEntry = b1TransferRequest.DocEntry;
                                b1Transfer.Lines.BaseLine = int.TryParse(line.ReferenceNumber, out var refKey) ? refKey : throw new Exception($"Cannot parse {line.ReferenceNumber}");
                                b1Transfer.Lines.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;
                                b1Transfer.Lines.Quantity = (double) line.ReceivedQuantity;

                                b1Transfer.Lines.Quantity /= GetNumPerMsr(b1TransferRequest.DocEntry, b1Transfer.Lines.BaseLine, "WTQ1");

                                var serialLine = 0;
                                if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                    foreach (var lineDetails in line.LineDetails)
                                    {
                                        if (serialLine != 0)
                                            b1Transfer.Lines.SerialNumbers.Add();
                                        b1Transfer.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                        b1Transfer.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(lineDetails.SerialNumber, line.Product.Sku);
                                        b1Transfer.Lines.SerialNumbers.InternalSerialNumber = lineDetails.SerialNumber;
                                    }

                                if (item.ManageBatchNumbers != BoYesNoEnum.tYES)
                                    continue;

                                var batchLine = 0;
                                foreach (var adjustment in line.LineDetails.GroupBy(c => c.LotNumber))
                                {
                                    if (batchLine != 0)
                                        b1Transfer.Lines.BatchNumbers.Add();
                                    b1Transfer.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                    b1Transfer.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                    b1Transfer.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1Transfer.Lines.BatchNumbers.ExpiryDate;
                                    b1Transfer.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.ReceivedQuantity * (c.PacksizeEachCount ?? 1));
                                }
                            }

                            StagingTableLockUnlock(b1TransferRequest.DocEntry.ToString(), (int)BoObjectTypes.oInventoryTransferRequest, false);
                            if (b1Transfer.Add() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"WH Transfer [{po.PurchaseOrderNumber}] upload failed: {errorCode} - {errorMessage}");
                            }
                            
                            WebInvoke("api/PurchaseOrderApi/CreateOrUpdate", null, Method.POST, new
                            {
                                po.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = true,
                                UploadMessage = (string) null,
                            });

                            StagingTableDelete(b1TransferRequest.DocEntry.ToString(), (int)BoObjectTypes.oInventoryTransferRequest);

                            Log($"WH Transfer: [{po.PurchaseOrderNumber}] for [{companySettings.ClientName ?? companySettings.CompanyDb}] uploaded");
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
                           
                        }
                        finally
                        {
                            if (item != null)
                                Marshal.ReleaseComObject(item);
                            if (b1Transfer != null)
                                Marshal.ReleaseComObject(b1Transfer);
                            if (b1TransferRequest != null)
                                Marshal.ReleaseComObject(b1TransferRequest);
                            Disconnect();
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
            }
        }
    }
}