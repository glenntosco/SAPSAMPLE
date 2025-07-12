using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using RestSharp;
using Pro4Soft.SapB1Integration.Infrastructure;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class WorkOrderUpload : IntegrationBase
    {
        public WorkOrderUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var wos = WebInvoke<List<WorkOrder>>("odata/WorkOrder", request =>
                    {
                        if (string.IsNullOrWhiteSpace(companySettings.ClientName))
                            request.AddQueryParameter("$filter", $@"UploadDate eq null and WorkOrderState eq 'Closed' and ClientId eq null");
                        else
                            request.AddQueryParameter("$filter", $@"UploadDate eq null and WorkOrderState eq 'Closed' and Client/Name eq '{companySettings.ClientName}'");
                        request.AddQueryParameter("$select", @"Id,WorkOrderNumber,ReferenceNumber,ProducedQuantity");
                        request.AddQueryParameter("$orderby", @"WorkOrderNumber");
                        request.AddQueryParameter("$expand", @"
Product,ProductDetails,WorkAreaBin,AssignedUser($select=Id,Username),Warehouse($select=Id,WarehouseCode),
Lines($orderby=LineNumber;$expand=Product,ConsumedProductDetails)");
                    }, Method.GET, null, "value");
                    if (!wos.Any())
                        continue;

                    foreach (var wo in wos)
                    {
                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
                        var b1GoodsReceipt = CurrentCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry) as Documents;
                        var b1GoodsIssue = CurrentCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit) as Documents;
                        try
                        {
                            if (!int.TryParse(wo.ReferenceNumber, out var woKey))
                                throw new Exception($"Cannot parse {wo.ReferenceNumber}");
                            int serialLine = 0;

                            if (!IsWoReceiptUploaded(woKey))
                            {
                                //Receipt
                                b1GoodsReceipt.DocDate = DateTime.Now;

                                b1GoodsReceipt.Lines.SetCurrentLine(0);
                                b1GoodsReceipt.Lines.BaseEntry = woKey;
                                b1GoodsReceipt.Lines.BaseType = (int) BoObjectTypes.oProductionOrders;
                                b1GoodsReceipt.Lines.Quantity = (double) wo.ProducedQuantity;

                                if (!item.GetByKey(wo.Product.Sku))
                                    throw new Exception($"GetByKey Failed for item: {wo.Product}");

                                if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                    foreach (var productDetails in wo.ProductDetails)
                                    {
                                        if (serialLine != 0)
                                            b1GoodsReceipt.Lines.SerialNumbers.Add();
                                        b1GoodsReceipt.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                        b1GoodsReceipt.Lines.SerialNumbers.InternalSerialNumber = productDetails.SerialNumber;
                                    }

                                if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                                {
                                    var batchLine = 0;
                                    foreach (var productDetails in wo.ProductDetails.GroupBy(c => c.LotNumber))
                                    {
                                        if (batchLine != 0)
                                            b1GoodsReceipt.Lines.BatchNumbers.Add();
                                        b1GoodsReceipt.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                        b1GoodsReceipt.Lines.BatchNumbers.BatchNumber = productDetails.Key;
                                        b1GoodsReceipt.Lines.BatchNumbers.ExpiryDate = productDetails.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1GoodsReceipt.Lines.BatchNumbers.ExpiryDate;
                                        b1GoodsReceipt.Lines.BatchNumbers.Quantity = (double) productDetails.Sum(c => c.ProducedQuantity * (c.PacksizeEachCount ?? 1));
                                    }
                                }

                                if (b1GoodsReceipt.Add() != 0)
                                {
                                    CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                    throw new Exception($"WO [{wo.WorkOrderNumber}] Product Receipt upload failed: {errorCode} - {errorMessage}");
                                }

                                StagingSetWoReceiptUploaded(woKey);
                                Log($"WO: [{wo.WorkOrderNumber}] for [{companySettings.ClientName ?? companySettings.CompanyDb}] Production Receipt created");
                            }

                            //Issues
                            b1GoodsIssue.DocDate = DateTime.Now;
                            int issueLine = 0;
                            foreach (var line in wo.Lines)
                            {
                                if (issueLine != 0)
                                    b1GoodsIssue.Lines.Add();
                                b1GoodsIssue.Lines.SetCurrentLine(issueLine++);

                                b1GoodsIssue.Lines.BaseEntry = woKey;
                                b1GoodsIssue.Lines.BaseLine = int.TryParse(line.ReferenceNumber, out var woLineKey) ? woLineKey : throw new Exception($"Cannot parse {line.ReferenceNumber}");
                                b1GoodsIssue.Lines.BaseType = (int) BoObjectTypes.oProductionOrders;
                                b1GoodsIssue.Lines.Quantity = (double) line.ConsumedQuantity;

                                if (!item.GetByKey(line.Product.Sku))
                                    throw new Exception($"GetByKey Failed for item: {line.Product}");

                                serialLine = 0;
                                if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                    foreach (var productDetails in line.ConsumedProductDetails)
                                    {
                                        if (serialLine != 0)
                                            b1GoodsIssue.Lines.SerialNumbers.Add();
                                        b1GoodsIssue.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                        b1GoodsIssue.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(productDetails.SerialNumber, line.Product.Sku);
                                        b1GoodsIssue.Lines.SerialNumbers.InternalSerialNumber = productDetails.SerialNumber;
                                    }

                                if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                                {
                                    var batchLine = 0;
                                    foreach (var productDetails in line.ConsumedProductDetails.GroupBy(c => c.LotNumber))
                                    {
                                        if (batchLine != 0)
                                            b1GoodsIssue.Lines.BatchNumbers.Add();
                                        b1GoodsIssue.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                        b1GoodsIssue.Lines.BatchNumbers.BatchNumber = productDetails.Key;
                                        //b1GoodsIssue.Lines.BatchNumbers.ExpiryDate = productDetails.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1GoodsIssue.Lines.BatchNumbers.ExpiryDate;
                                        b1GoodsIssue.Lines.BatchNumbers.Quantity = (double) productDetails.Sum(c => c.ConsumedQuantity * (c.PacksizeEachCount ?? 1));
                                    }
                                }
                            }

                            if (b1GoodsIssue.Add() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"WO [{wo.WorkOrderNumber}] Product Issue upload failed: {errorCode} - {errorMessage}");
                            }

                            WebInvoke("api/WorkOrderApi/CreateOrUpdate", null, Method.POST, new
                            {
                                wo.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = true,
                                UploadMessage = (string) null,
                            });

                            StagingTableIncrementBackOrder(wo.ReferenceNumber, (int) BoObjectTypes.oProductionOrders);
                            Log($"WO: [{wo.WorkOrderNumber}] for [{companySettings.ClientName ?? companySettings.CompanyDb}] Production Issue created");
                        }
                        catch (Exception e)
                        {
                            WebInvoke("api/WorkOrderApi/CreateOrUpdate", null, Method.POST, new
                            {
                                wo.Id,
                                UploadDate = DateTime.UtcNow,
                                UploadedSuceeded = false,
                                UploadMessage = e.ToString()
                            });
                            LogError(e);
                            StagingTableLockUnlock(wo.ReferenceNumber, (int) BoObjectTypes.oProductionOrders, true);
                        }
                        finally
                        {
                            if (item != null)
                                Marshal.ReleaseComObject(item);
                            if (b1GoodsReceipt != null)
                                Marshal.ReleaseComObject(b1GoodsReceipt);
                            if (b1GoodsIssue != null)
                                Marshal.ReleaseComObject(b1GoodsIssue);
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