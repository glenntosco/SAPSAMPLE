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
    public class CustVendReturnsAdjustmentUpload : IntegrationBase
    {
        public CustVendReturnsAdjustmentUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            var adjusts = WebInvoke<List<ReturnAdjustment>>("api/Audit/Execute", null, Method.POST, new
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
                            Condition = "Or",
                            Rules = new List<dynamic>
                            {
                                new
                                {
                                    Field = "SubType",
                                    Value = "CustomerReturnAdjustIn",
                                    Operator = "Equal"
                                },
                                new
                                {
                                    Field = "SubType",
                                    Value = "VendorReturnAdjustOut",
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
                SetCompany(companySettings);
                try
                {
                    var clientAdjustments = adjusts.ToList();
                    if (!string.IsNullOrWhiteSpace(companySettings.ClientName))
                        clientAdjustments = adjusts.Where(c => c.Client == companySettings.ClientName).ToList();

                    var reference = Guid.NewGuid();
                    var currentDataSet = clientAdjustments.Where(c => c.ReferenceType == "Customer").ToList();
                    try
                    {
                        CreateCustomerReturn(currentDataSet, reference, companySettings);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = (string)null
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

                    currentDataSet = clientAdjustments.Where(c => c.ReferenceType == "Vendor").ToList();
                    try
                    {
                        CreateVendorReturn(currentDataSet, reference, companySettings);
                        WebInvoke("api/Audit/SetIntegrationReference", null, Method.POST, new
                        {
                            Ids = currentDataSet.Select(c => c.Id).ToList(),
                            IntegrationReference = reference,
                            IntegrationMessage = (string)null
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
                catch (Exception ex)
                {
                    LogError(ex);
                }
                finally
                {
                    Disconnect();
                }
            }
        }

        private void CreateCustomerReturn(List<ReturnAdjustment> clientAdjustments, Guid reference, CompanySettings companySettings)
        {
            var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;

            foreach (var customerAdjusts in clientAdjustments.GroupBy(c => c.ReferenceId))
            {
                var b1CustomerReturn = CurrentCompany.GetBusinessObject(BoObjectTypes.oReturns) as Documents;
                try
                {
                    var customer = WebInvoke<List<dynamic>>("odata/Customer", request =>
                    {
                        request.AddQueryParameter("$filter", $"Id eq {customerAdjusts.Key}");
                        request.AddQueryParameter("$select", "CustomerCode");
                    }, Method.GET, null, "value");
                    string customerCode = customer.SingleOrDefault()?.CustomerCode.ToString();
                    if (string.IsNullOrWhiteSpace(customerCode))
                    {
                        LogError($"Customer {customerAdjusts.Key} does not exist");
                        continue;
                    }

                    b1CustomerReturn.DocDate = DateTime.Now;
                    b1CustomerReturn.TaxDate = DateTime.Now;
                    b1CustomerReturn.Comments = reference.ToString();
                    b1CustomerReturn.CardCode = customerCode;
                    var lineNum = 0;

                    foreach (var skuAdjust in customerAdjusts.GroupBy(c => c.Sku))
                    {
                        if (!item.GetByKey(skuAdjust.Key))
                            throw new Exception($"Unable to retrieve item: {skuAdjust.Key}");
                        foreach (var skuAdjustToWarehouse in skuAdjust.GroupBy(c => c.ToWarehouse))
                        {
                            if (lineNum != 0)
                                b1CustomerReturn.Lines.Add();
                            b1CustomerReturn.Lines.SetCurrentLine(lineNum++);
                            b1CustomerReturn.Lines.ItemCode = skuAdjust.Key;
                            b1CustomerReturn.Lines.Quantity = (double) skuAdjustToWarehouse.Sum(c => c.Quantity);
                            b1CustomerReturn.Lines.WarehouseCode = skuAdjustToWarehouse.Select(c => c.ToWarehouse).FirstOrDefault();

                            var price = GetAdjustmentPrice(item.ItemCode);
                            if (price != null)
                                b1CustomerReturn.Lines.Price = price.Value;


                            if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                            {
                                var batchLine = 0;
                                foreach (var adjustment in skuAdjustToWarehouse.GroupBy(c => c.LotNumber))
                                {
                                    if (batchLine != 0)
                                        b1CustomerReturn.Lines.BatchNumbers.Add();
                                    b1CustomerReturn.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                    b1CustomerReturn.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                    b1CustomerReturn.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1CustomerReturn.Lines.BatchNumbers.ExpiryDate;
                                    b1CustomerReturn.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.Quantity * (c.EachCount ?? 1));
                                }
                            }

                            if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                            {
                                var serialLine = 0;
                                foreach (var adjustment in skuAdjustToWarehouse)
                                {
                                    if (serialLine != 0)
                                        b1CustomerReturn.Lines.SerialNumbers.Add();
                                    b1CustomerReturn.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                    b1CustomerReturn.Lines.SerialNumbers.InternalSerialNumber = adjustment.SerialNumber;
                                    b1CustomerReturn.Lines.SerialNumbers.Quantity = 1;
                                }
                            }
                        }
                    }

                    if (lineNum <= 0)
                        continue;

                    if (b1CustomerReturn.Add() != 0)
                    {
                        CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                        throw new Exception($"Customer Return failed: {errorCode} - {errorMessage}");
                    }

                    Log($"Customer Return for [{companySettings.ClientName ?? companySettings.CompanyDb}] created");
                }
                catch (Exception e)
                {
                    LogError(e);
                }
                finally
                {
                    if (b1CustomerReturn != null)
                        Marshal.ReleaseComObject(b1CustomerReturn);
                }
            }

            if (item != null)
                Marshal.ReleaseComObject(item);
        }

        private void CreateVendorReturn(List<ReturnAdjustment> clientAdjustments, Guid reference, CompanySettings companySettings)
        {
            var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
            foreach (var vendorAdjusts in clientAdjustments.GroupBy(c => c.ReferenceId))
            {
                var b1VendorReturn = CurrentCompany.GetBusinessObject(BoObjectTypes.oPurchaseReturns) as Documents;
                try
                {
                    var vendor = WebInvoke<List<dynamic>>("odata/Vendor", request =>
                    {
                        request.AddQueryParameter("$filter", $"Id eq {vendorAdjusts.Key}");
                        request.AddQueryParameter("$select", "VendorCode");
                    }, Method.GET, null, "value");
                    string vendorCode = vendor.SingleOrDefault()?.VendorCode.ToString();
                    if (string.IsNullOrWhiteSpace(vendorCode))
                        throw new BusinessWebException($"Vendor {vendorAdjusts.Key} does not exist");

                    b1VendorReturn.DocDate = DateTime.Now;
                    b1VendorReturn.TaxDate = DateTime.Now;
                    b1VendorReturn.Comments = reference.ToString();
                    b1VendorReturn.CardCode = vendorCode;
                    var lineNum = 0;

                    foreach (var skuAdjust in vendorAdjusts.GroupBy(c => c.Sku))
                    {
                        if (!item.GetByKey(skuAdjust.Key))
                            throw new Exception($"Unable to retrieve item: {skuAdjust.Key}");
                        foreach (var skuAdjustFromWarehouse in skuAdjust.GroupBy(c => c.FromWarehouse))
                        {
                            if (lineNum != 0)
                                b1VendorReturn.Lines.Add();
                            b1VendorReturn.Lines.SetCurrentLine(lineNum++);
                            b1VendorReturn.Lines.ItemCode = skuAdjust.Key;
                            b1VendorReturn.Lines.Quantity = (double) skuAdjustFromWarehouse.Sum(c => c.Quantity);
                            b1VendorReturn.Lines.WarehouseCode = skuAdjustFromWarehouse.Select(c => c.FromWarehouse).FirstOrDefault();

                            var price = GetAdjustmentPrice(item.ItemCode);
                            if (price != null)
                                b1VendorReturn.Lines.Price = price.Value;

                            if (item.ManageBatchNumbers == BoYesNoEnum.tYES)
                            {
                                var batchLine = 0;
                                foreach (var adjustment in skuAdjustFromWarehouse.GroupBy(c => c.LotNumber))
                                {
                                    if (batchLine != 0)
                                        b1VendorReturn.Lines.BatchNumbers.Add();
                                    b1VendorReturn.Lines.BatchNumbers.SetCurrentLine(batchLine++);
                                    b1VendorReturn.Lines.BatchNumbers.BatchNumber = adjustment.Key;
                                    b1VendorReturn.Lines.BatchNumbers.ExpiryDate = adjustment.FirstOrDefault()?.ExpiryDate?.DateTime ?? b1VendorReturn.Lines.BatchNumbers.ExpiryDate;
                                    b1VendorReturn.Lines.BatchNumbers.Quantity = (double) adjustment.Sum(c => c.Quantity * (c.EachCount ?? 1));
                                }
                            }

                            if (item.ManageSerialNumbers == BoYesNoEnum.tYES)
                            {
                                var serialLine = 0;
                                foreach (var adjustment in skuAdjustFromWarehouse)
                                {
                                    if (serialLine != 0)
                                        b1VendorReturn.Lines.SerialNumbers.Add();
                                    b1VendorReturn.Lines.SerialNumbers.SetCurrentLine(serialLine++);
                                    b1VendorReturn.Lines.SerialNumbers.SystemSerialNumber = GetSystemSerial(adjustment.SerialNumber, adjustment.Sku);
                                    b1VendorReturn.Lines.SerialNumbers.InternalSerialNumber = adjustment.SerialNumber;
                                    b1VendorReturn.Lines.SerialNumbers.Quantity = 1;
                                }
                            }
                        }
                    }

                    if (lineNum <= 0)
                        continue;

                    if (b1VendorReturn.Add() != 0)
                    {
                        CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                        throw new Exception($"Vendor Return failed: {errorCode} - {errorMessage}");
                    }

                    Log($"Vendor Return for [{companySettings.ClientName ?? companySettings.CompanyDb}] created");
                }
                catch (Exception e)
                {
                    LogError(e);
                }
                finally
                {
                    if (b1VendorReturn != null)
                        Marshal.ReleaseComObject(b1VendorReturn);
                }
            }
            if (item != null)
                Marshal.ReleaseComObject(item);
        }
    }
}