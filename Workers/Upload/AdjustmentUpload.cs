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
    public class AdjustmentUpload : IntegrationBase
    {
        public AdjustmentUpload(ScheduleSetting settings) : base(settings)
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
                            Field = "ReferenceType",
                            Value = "Invoice",
                            Operator = "NotEqual"
                        },
                        new
                        {
                            Condition = "Or",
                            Rules = new List<dynamic>
                            {
                                new
                                {
                                    Field = "SubType",
                                    Value = "AdjustIn",
                                    Operator = "Equal"
                                },
                                new
                                {
                                    Field = "SubType",
                                    Value = "AdjustOut",
                                    Operator = "Equal"
                                }
                            }
                        }
                    }
                }
            });
            if (!adjusts.Any())
                return;
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var clientAdjustments = adjusts.ToList();
                    if (!string.IsNullOrWhiteSpace(companySettings.ClientName))
                        clientAdjustments = adjusts.Where(c => c.Client == companySettings.ClientName).ToList();

                    var reference = Guid.NewGuid();
                    var currentDataSet = clientAdjustments.Where(c => c.SubType == "AdjustIn").ToList();
                    try
                    {
                        CreateGoodsReceipt(currentDataSet, reference, companySettings);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = (string) null
                        });
                    }
                    catch (Exception e)
                    {
                        LogError(e);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = e.Message
                        });
                    }

                    currentDataSet = clientAdjustments.Where(c => c.SubType == "AdjustOut").ToList();
                    try
                    {
                        CreateGoodsIssue(currentDataSet, reference, companySettings);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = (string) null
                        });
                    }
                    catch (Exception e)
                    {
                        LogError(e);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = e.Message
                        });
                    }
                }
                finally
                {
                    Disconnect();
                }
            }
        }

        private void CreateGoodsReceipt(IEnumerable<Adjustment> clientAdjustments, Guid reference, CompanySettings companySettings)
        {
            var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
            var b1GoodsReceipt = CurrentCompany.GetBusinessObject(BoObjectTypes.oInventoryGenEntry) as Documents;

            try
            {
                b1GoodsReceipt.DocDate = DateTime.Now;
                b1GoodsReceipt.TaxDate = DateTime.Now;
                b1GoodsReceipt.Comments = reference.ToString();
                b1GoodsReceipt.Reference2 = "Pro4Soft";
                var lineNum = 0;

                //Inbound
                foreach (var whAdjusts in clientAdjustments.GroupBy(c => c.ToWarehouse))
                {
                    foreach (var skuAdjusts in whAdjusts.GroupBy(c => c.Sku))
                    {
                        if (!item.GetByKey(skuAdjusts.Key))
                            throw new Exception($"Unable to retrieve item: {skuAdjusts.Key}");

                        if (lineNum != 0)
                            b1GoodsReceipt.Lines.Add();

                        b1GoodsReceipt.Lines.SetCurrentLine(lineNum++);
                        b1GoodsReceipt.Lines.ItemCode = skuAdjusts.Key;
                        b1GoodsReceipt.Lines.Quantity = (double) skuAdjusts.Sum(c => c.Quantity);
                        b1GoodsReceipt.Lines.WarehouseCode = whAdjusts.Key;

                        var price = GetAdjustmentPrice(item.ItemCode);
                        if (price != null)
                            b1GoodsReceipt.Lines.Price = price.Value;

                        if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                        {
                            var batchLine = 0;
                            foreach (var adjustment in skuAdjusts.GroupBy(c => c.LotNumber))
                            {
                                if (batchLine != 0)
                                    b1GoodsReceipt.Lines.BatchNumbers.Add();
                                b1GoodsReceipt.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                b1GoodsReceipt.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                b1GoodsReceipt.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1GoodsReceipt.Lines.BatchNumbers.ExpiryDate;
                                b1GoodsReceipt.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.Quantity * (c.EachCount ?? 1));
                            }
                        }

                        if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                        {
                            var serialLine = 0;
                            foreach (var adjustment in skuAdjusts)
                            {
                                if (serialLine != 0)
                                    b1GoodsReceipt.Lines.SerialNumbers.Add();
                                b1GoodsReceipt.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                b1GoodsReceipt.Lines.SerialNumbers.InternalSerialNumber = adjustment.SerialNumber;
                                b1GoodsReceipt.Lines.SerialNumbers.Quantity = 1;
                            }
                        }
                    }
                }

                if (lineNum <= 0)
                    return;

                if (b1GoodsReceipt.Add() != 0)
                {
                    CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                    throw new Exception($"Goods Receipt failed: {errorCode} - {errorMessage}");
                }

                Log($"Goods Receipt for [{companySettings.ClientName ?? companySettings.CompanyDb}] created");
            }
            finally
            {
                if(item != null)
                    Marshal.ReleaseComObject(item);
                if (b1GoodsReceipt != null)
                    Marshal.ReleaseComObject(b1GoodsReceipt);
            }
        }

        private void CreateGoodsIssue(IEnumerable<Adjustment> clientAdjustments, Guid reference, CompanySettings companySettings)
        {
            var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
            var b1GoodsIssue = CurrentCompany.GetBusinessObject(BoObjectTypes.oInventoryGenExit) as Documents;
            try
            {
                b1GoodsIssue.DocDate = DateTime.Now;
                b1GoodsIssue.TaxDate = DateTime.Now;
                b1GoodsIssue.Comments = reference.ToString();
                b1GoodsIssue.Reference2 = "Pro4Soft";
                var lineNum = 0;

                //Outbound
                foreach (var whAdjusts in clientAdjustments.Where(c => c.SubType == "AdjustOut").GroupBy(c => c.FromWarehouse))
                {
                    foreach (var skuAdjusts in whAdjusts.GroupBy(c => c.Sku))
                    {
                        if (!item.GetByKey(skuAdjusts.Key))
                            throw new Exception($"Unable to retrieve item: {skuAdjusts.Key}");

                        if (lineNum != 0)
                            b1GoodsIssue.Lines.Add();
                        b1GoodsIssue.Lines.SetCurrentLine(lineNum++);
                        b1GoodsIssue.Lines.ItemCode = skuAdjusts.Key;
                        b1GoodsIssue.Lines.Quantity = (double)skuAdjusts.Sum(c => c.Quantity);
                        b1GoodsIssue.Lines.WarehouseCode = whAdjusts.Key;

                        var price = GetAdjustmentPrice(item.ItemCode);
                        if (price != null)
                            b1GoodsIssue.Lines.Price = price.Value;

                        if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                        {
                            var batchLine = 0;
                            foreach (var adjustment in skuAdjusts.GroupBy(c => c.LotNumber))
                            {
                                if (batchLine != 0)
                                    b1GoodsIssue.Lines.BatchNumbers.Add();
                                b1GoodsIssue.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                b1GoodsIssue.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                b1GoodsIssue.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1GoodsIssue.Lines.BatchNumbers.ExpiryDate;
                                b1GoodsIssue.Lines.BatchNumbers.Quantity = (double)adjustment.Sum(c => c.Quantity * (c.EachCount ?? 1));
                            }
                        }

                        if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                        {
                            var serialLine = 0;
                            foreach (var adjustment in skuAdjusts)
                            {
                                if (serialLine != 0)
                                    b1GoodsIssue.Lines.SerialNumbers.Add();
                                b1GoodsIssue.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                b1GoodsIssue.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(adjustment.SerialNumber, adjustment.Sku);
                                b1GoodsIssue.Lines.SerialNumbers.InternalSerialNumber = adjustment.SerialNumber;
                                b1GoodsIssue.Lines.SerialNumbers.Quantity = 1;
                            }
                        }
                    }
                }

                if (lineNum <= 0)
                    return;

                if (b1GoodsIssue.Add() != 0)
                {
                    CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                    throw new Exception($"Goods Issue failed: {errorCode} - {errorMessage}");
                }
                Log($"Goods Issue for [{companySettings.ClientName ?? companySettings.CompanyDb}] created");
            }
            finally
            {
                if (item != null)
                    Marshal.ReleaseComObject(item);
                if (b1GoodsIssue != null)
                    Marshal.ReleaseComObject(b1GoodsIssue);
            }
        }
    }

    public enum AdjustmentsPricingSchema
    {
        Default,
        FixedPriceList,
        LastEvaluatedPrice,
        LastPurchasePrice
    }
}