using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;
// ReSharper disable InconsistentNaming

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class SimpleWarehouseTransferUpload : IntegrationBase
    {
        public SimpleWarehouseTransferUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            var adjusts = WebInvoke<List<Adjustment>>("api/Audit/Execute", null, Method.POST, new
            {
                Filter = new
                {
                    Condition = "And",
                    Rules = new List<dynamic>
                    {
                        new
                        {
                            Field = "IntegrationReference",
                            Operator = "IsNull"
                        },
                        new
                        {
                            Field = "SubType",
                            Value = "WarehouseMove",
                            Operator = "Equal"
                        }
                    }
                }
            });
            if (!adjusts.Any())
                return;
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                SetCompany(companySettings);
                var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
                try
                {
                    var adjustments = adjusts.ToList();
                    if (!string.IsNullOrWhiteSpace(companySettings.ClientName))
                        adjustments = adjusts.Where(c => c.Client == companySettings.ClientName).ToList();

                    foreach (var transferFromWarehouse in adjustments.GroupBy(c => c.FromWarehouse))
                    {
                        var reference = Guid.NewGuid();
                        var b1StockTransfer = CurrentCompany.GetBusinessObject(BoObjectTypes.oStockTransfer) as StockTransfer;
                        try
                        {
                            var transferLinesCount = 0;
                            b1StockTransfer.FromWarehouse = transferFromWarehouse.Key;
                            b1StockTransfer.ToWarehouse = transferFromWarehouse.First().ToWarehouse;
                            b1StockTransfer.TaxDate = DateTime.Now;
                            b1StockTransfer.DocDate = DateTime.Now;
                            b1StockTransfer.Comments = reference.ToString();

                            foreach (var skuAdjust in transferFromWarehouse.GroupBy(c => c.Sku))
                            {
                                if (!item.GetByKey(skuAdjust.Key))
                                    throw new Exception($"Unable to retrieve item: {skuAdjust.Key}");
                                foreach (var skuAdjustToWarehouse in skuAdjust.GroupBy(c => c.ToWarehouse))
                                {
                                    if (transferLinesCount != 0)
                                        b1StockTransfer.Lines.Add();
                                    b1StockTransfer.Lines.SetCurrentLine(transferLinesCount++);

                                    b1StockTransfer.Lines.ItemCode = skuAdjust.Key;
                                    b1StockTransfer.Lines.WarehouseCode = skuAdjustToWarehouse.Key;
                                    b1StockTransfer.Lines.Quantity = (double) skuAdjustToWarehouse.Sum(c => c.Quantity);

                                    if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                                    {
                                        var batchLine = 0;
                                        foreach (var adjustment in skuAdjustToWarehouse.GroupBy(c => c.LotNumber))
                                        {
                                            if (batchLine != 0)
                                                b1StockTransfer.Lines.BatchNumbers.Add();
                                            b1StockTransfer.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                            b1StockTransfer.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                            b1StockTransfer.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1StockTransfer.Lines.BatchNumbers.ExpiryDate;
                                            b1StockTransfer.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.Quantity * (c.EachCount ?? 1));
                                        }
                                    }

                                    if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                                    {
                                        var serialLine = 0;
                                        foreach (var adjustment in skuAdjustToWarehouse)
                                        {
                                            if (serialLine != 0)
                                                b1StockTransfer.Lines.SerialNumbers.Add();
                                            b1StockTransfer.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                            b1StockTransfer.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(adjustment.SerialNumber, adjustment.Sku);
                                            b1StockTransfer.Lines.SerialNumbers.InternalSerialNumber = adjustment.SerialNumber;
                                            b1StockTransfer.Lines.SerialNumbers.Quantity = 1;
                                        }
                                    }
                                }
                            }

                            if (transferLinesCount > 0 && b1StockTransfer.Add() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"WH Transfer upload failed: {errorCode} - {errorMessage}");
                            }

                            WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                            {
                                Ids = transferFromWarehouse.Select(c => c.Id).ToList(),
                                IntegrationReference = reference,
                                IntegrationMessage = (string) null
                            });
                        }
                        catch (Exception e)
                        {
                            WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                            {
                                Ids = transferFromWarehouse.Select(c => c.Id).ToList(),
                                IntegrationReference = reference,
                                IntegrationMessage = e.Message
                            });
                        }
                        finally
                        {
                            if (b1StockTransfer != null)
                                Marshal.ReleaseComObject(b1StockTransfer);
                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
                finally
                {
                    if (item != null)
                        Marshal.ReleaseComObject(item);
                    Disconnect();
                }
            }
        }
    }
}