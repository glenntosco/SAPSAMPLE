using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class ItemDimensionsUpload : IntegrationBase
    {
        private static bool _isPrimed;
        private const string ItemDimensionsUploadLastTime = nameof(ItemDimensionsUploadLastTime);

        public ItemDimensionsUpload(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            Subscribe();

            var now = DateTimeOffset.UtcNow;
            var serverVal = WebInvoke<string>("api/TenantApi/GetTenantSettings", request => { request.AddQueryParameter("key", ItemDimensionsUploadLastTime); });
            var lastTime = DateTimeOffset.TryParse(serverVal, out var date) ? date : now.Subtract(TimeSpan.FromDays(60));

            var historyRecords = WebInvoke<List<ObjectChangeEvent>>("api/Audit/ExecuteHistory", null, Method.POST, new
            {
                Condition = "And",
                Rules = new List<dynamic>
                {
                    new
                    {
                        Field = "Timestamp",
                        Value = lastTime.ToString("O"),
                        Operator = "GreaterOrEqual"
                    },
                    new
                    {
                        Condition = "Or",
                        Rules = new List<dynamic>
                        {
                            new
                            {
                                Field = "ObjectType",
                                Value = "Product",
                                Operator = "Equal"
                            },
                            new
                            {
                                Field = "ObjectType",
                                Value = "ProductPacksize",
                                Operator = "Equal"
                            }
                        }
                    }
                }
            });
            if (!historyRecords.Any())
                return;

            var interestedFields = new[] {"Height", "Width", "Length", "Weight", "Upc", "BarcodeValue"};
            foreach (var objectTypeGroup in historyRecords
                .Where(c => c.Changes.Any(c1 => interestedFields.Contains(c1.PropertyName)))
                .GroupBy(c => c.ObjectType))
            {
                var distinctObjs = objectTypeGroup.Select(c => c.ObjectId).Distinct().ToList();
                switch (objectTypeGroup.Key)
                {
                    case "Product":
                    {
                        var prods = new List<FullProduct>();
                        for (var i = 0; i < distinctObjs.Count; i += 10)
                        {
                            var batch = distinctObjs.Skip(i).Take(10);
                            prods.AddRange(WebInvoke<List<FullProduct>>("odata/Product", request =>
                            {
                                request.AddQueryParameter("$filter", string.Join(" or ", batch.Select(c => $"(Id eq {c})")));
                                request.AddQueryParameter("$expand", "Client");
                            }, Method.GET, null, "value").ToList());
                        }
                        UpdateProductDims(prods);
                        break;
                    }
                    case "ProductPacksize":
                    {
                        var packsizes = new List<ProductPacksize>();
                        for (var i = 0; i < distinctObjs.Count; i += 10)
                        {
                            var batch = distinctObjs.Skip(i).Take(10);
                            packsizes.AddRange(WebInvoke<List<ProductPacksize>>("odata/ProductPacksize", request =>
                            {
                                request.AddQueryParameter("$filter", string.Join(" or ", batch.Select(c => $"(Id eq {c})")));
                                request.AddQueryParameter("$expand", "Product($expand=Client)");
                            }, Method.GET, null, "value").ToList());
                        }
                        UpdateProductPacksizeDims(packsizes);
                        break;
                    }
                }
            }

            WebInvoke("api/TenantApi/SetTenantSettings", request => { request.AddQueryParameter("key", ItemDimensionsUploadLastTime); }, Method.POST, now.ToString());
        }

        private void UpdateProductDims(List<FullProduct> products)
        {
            var companies = products.GroupBy(c => c.ClientId);
            foreach (var compProds in companies)
            {
                var first = compProds.First();
                var companySettings = App<SapConnectionSettings>.Instance.Companies.SingleOrDefault(c => c.ClientName == first.Client?.Name);
                if (companySettings == null)
                    continue;

                SetCompany(companySettings);
                foreach (var product in compProds)
                {
                    try
                    {
                        //Dims are driven based off of packsizes
                        if (product.IsPacksizeControlled)
                            continue;

                        var b1Product = ExecuteReader($@"
select
    OITM.CodeBars,
    OITM.BLength1,
    OITM.BHeight1, 
    OITM.BWeight1, 
    OITM.BWidth1
from 
    OITM
where 
    OITM.ItemCode = '{product.Sku}'").SingleOrDefault();
                        if (b1Product == null)
                            continue;

                        if (b1Product.BLength1 == (product.Length ?? 0) &&
                            b1Product.BWidth1 == (product.Width ?? 0) &&
                            b1Product.BHeight1 == (product.Height ?? 0) &&
                            b1Product.BWeight1 == (product.Weight ?? 0) &&
                            b1Product.BcdCode == product.Upc)
                            continue;

                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;

                        try
                        {
                            if (!item.GetByKey(product.Sku))
                                throw new Exception($"Unable to retrieve item: {product.Sku}");
                            var weightUnits = ExecuteScalar<int>($"select top 1 UnitCode from owgt where UnitDisply = '{product.DimsWeightUnitOfMeasure}'");
                            if (weightUnits == 0)
                                throw new Exception($"Units [{product.DimsWeightUnitOfMeasure}] are not setup in SAP (OWGT table)");
                            var dimensionUnits = ExecuteScalar<int>($"select top 1 UnitCode from olgt where UnitDisply = '{product.DimsLengthUnitOfMeasure}'");
                            if (dimensionUnits == 0)
                                throw new Exception($"Units [{product.DimsLengthUnitOfMeasure}] are not setup in SAP (OLGT) table");

                            if (item.BarCode != product.Upc)
                                item.BarCode = product.Upc;
                            if (Math.Abs(item.PurchaseUnitLength - (double) (product.Length ?? 0)) > 0.0001)
                            {
                                item.PurchaseUnitLength = (double) (product.Length ?? 0);
                                item.PurchaseLengthUnit = dimensionUnits;
                            }

                            if (Math.Abs(item.PurchaseUnitWidth - (double) (product.Width ?? 0)) > 0.0001)
                            {
                                item.PurchaseUnitWidth = (double) (product.Width ?? 0);
                                item.PurchaseWidthUnit = dimensionUnits;
                            }

                            if (Math.Abs(item.PurchaseUnitHeight - (double) (product.Height ?? 0)) > 0.0001)
                            {
                                item.PurchaseUnitHeight = (double) (product.Height ?? 0);
                                item.PurchaseHeightUnit = dimensionUnits;
                            }

                            if (Math.Abs(item.PurchaseUnitWeight - (double) (product.Weight ?? 0)) > 0.0001)
                            {
                                item.PurchaseUnitWeight = (double) (product.Weight ?? 0);
                                item.PurchaseWeightUnit = weightUnits;
                            }

                            if (item.Update() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"Failed to update dimensions for [{product.Sku}]: {errorCode} - {errorMessage}");
                            }

                            Log($"Sku: [{product.Sku}] dimensions uploaded");
                        }
                        finally
                        {
                            if (item != null)
                                Marshal.ReleaseComObject(item);
                        }
                    }
                    catch (Exception e)
                    {
                        LogError(e);
                    }
                }
                Disconnect();
            }
        }

        private void UpdateProductPacksizeDims(List<ProductPacksize> productPacksizes)
        {
            var companies = productPacksizes.GroupBy(c => c.Product.ClientId);
            foreach (var compPacksizes in companies)
            {
                var first = compPacksizes.First();
                var companySettings = App<SapConnectionSettings>.Instance.Companies.SingleOrDefault(c => c.ClientName == first.Product.Client?.Name);
                if (companySettings == null)
                    continue;

                SetCompany(companySettings);
                foreach (var packsize in compPacksizes)
                {
                    try
                    {
                        var b1Packsize = ExecuteReader($@"
select
    OUOM.UomEntry,
    OBCD.BcdCode,
    ITM12.Length1, 
    ITM12.Width1, 
    ITM12.Height1,
	ITM12.Weight1    
from 
    OITM	
    left outer join UGP1 on UGP1.UgpEntry = OITM.UgpEntry
    left outer join OUOM on UGP1.UomEntry = OUOM.UomEntry
    left outer join ITM12 on ITM12.ItemCode = OITM.ItemCode and OUOM.UomEntry = ITM12.UomEntry and ITM12.UomType = 'P'
    left outer join OBCD on OBCD.ItemCode = OITM.ItemCode and OUOM.UomEntry = OBCD.UomEntry
where
	OITM.ItemCode = '{packsize.Product.Sku}' and 
	(UGP1.BaseQty/UGP1.AltQty) = {packsize.EachCount}").SingleOrDefault();
                        if (b1Packsize == null)
                            continue;

                        if (b1Packsize.Length1 == (packsize.Length ?? 0) &&
                            b1Packsize.Width1 == (packsize.Width ?? 0) &&
                            b1Packsize.Height1 == (packsize.Height ?? 0) &&
                            b1Packsize.Weight1 == (packsize.Weight ?? 0) &&
                            b1Packsize.BcdCode == packsize.BarcodeValue)
                            continue;

                        var item = CurrentCompany.GetBusinessObject(BoObjectTypes.oItems) as Items;
                        try
                        {
                            if (!item.GetByKey(packsize.Product.Sku))
                                throw new Exception($"Unable to retrieve item: {packsize.Product.Sku}");

                            var weightUnits = ExecuteScalar<int>($"select top 1 UnitCode from owgt where UnitDisply = '{packsize.Product.DimsWeightUnitOfMeasure}'");
                            if (weightUnits == 0)
                                throw new Exception($"Units [{packsize.Product.DimsWeightUnitOfMeasure}] are not setup in SAP (OWGT table)");
                            var dimensionUnits = ExecuteScalar<int>($"select top 1 UnitCode from olgt where UnitDisply = '{packsize.Product.DimsLengthUnitOfMeasure}'");
                            if (dimensionUnits == 0)
                                throw new Exception($"Units [{packsize.Product.DimsLengthUnitOfMeasure}] are not setup in SAP (OLGT) table");

                            var found = false;
                            for (var i = 0; i < item.UnitOfMeasurements.Count; i++)
                            {
                                item.UnitOfMeasurements.SetCurrentLine(i);
                                if (item.UnitOfMeasurements.UoMEntry != b1Packsize.UomEntry || item.UnitOfMeasurements.UoMType != ItemUoMTypeEnum.iutPurchasing)
                                    continue;
                                found = true;
                                break;
                            }

                            if (found)
                            {
                                if (Math.Abs(item.UnitOfMeasurements.Length1 - (double) (packsize.Length ?? 0)) > 0.0001)
                                {
                                    item.UnitOfMeasurements.Length1 = (double) (packsize.Length ?? 0);
                                    item.UnitOfMeasurements.Length1Unit = dimensionUnits;
                                }

                                if (Math.Abs(item.UnitOfMeasurements.Width1 - (double) (packsize.Width ?? 0)) > 0.0001)
                                {
                                    item.UnitOfMeasurements.Width1 = (double) (packsize.Width ?? 0);
                                    item.UnitOfMeasurements.Width1Unit = dimensionUnits;
                                }

                                if (Math.Abs(item.UnitOfMeasurements.Height1 - (double) (packsize.Height ?? 0)) > 0.0001)
                                {
                                    item.UnitOfMeasurements.Height1 = (double) (packsize.Height ?? 0);
                                    item.UnitOfMeasurements.Height1Unit = dimensionUnits;
                                }

                                if (Math.Abs(item.UnitOfMeasurements.Weight1 - (double) (packsize.Weight ?? 0)) > 0.0001)
                                {
                                    item.UnitOfMeasurements.Weight1 = (double) (packsize.Weight ?? 0);
                                    item.UnitOfMeasurements.Weight1Unit = weightUnits;
                                }
                            }

                            found = false;
                            for (var i = 0; i < item.BarCodes.Count; i++)
                            {
                                item.BarCodes.SetCurrentLine(i);
                                if (item.BarCodes.UoMEntry != b1Packsize.UomEntry)
                                    continue;
                                found = true;
                                break;
                            }

                            if (!found)
                            {
                                if (item.BarCodes.Count > 1)
                                {
                                    item.BarCodes.Add();
                                    item.BarCodes.SetCurrentLine(item.BarCodes.Count - 1);
                                }

                                item.BarCodes.UoMEntry = b1Packsize.UomEntry;
                                item.BarCodes.BarCode = packsize.BarcodeValue;
                            }
                            else if (item.BarCodes.BarCode != packsize.BarcodeValue)
                                item.BarCodes.BarCode = packsize.BarcodeValue;


                            if (item.Update() != 0)
                            {
                                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                                throw new Exception($"Failed to update dimensions for packsize [{packsize.Name}] of [{packsize.Product.Sku}]: {errorCode} - {errorMessage}");
                            }

                            Log($"Packsize [{packsize.Name}] of [{packsize.Product.Sku}] dimensions uploaded");
                        }
                        finally
                        {
                            if (item != null)
                                Marshal.ReleaseComObject(item);
                        }
                    }
                    catch (Exception e)
                    {
                        LogError(e);
                    }
                }
                Disconnect();
            }
        }

        private void Subscribe()
        {
            if (_isPrimed)
                return;

            Singleton<WebEventListener>.Instance.Subscribe("ObjectChanged", p =>
            {
                try
                {
                    ObjectChangeEvent changeEvent = Utils.DeserializeFromJson<ObjectChangeEvent>(Utils.SerializeToStringJson(p));
                    switch (changeEvent.ObjectType)
                    {
                        case "Product" when changeEvent.Changes.Any(c => c.PropertyName == "Height" || c.PropertyName == "Width" || c.PropertyName == "Length" || c.PropertyName == "Weight" || c.PropertyName == "Upc"):
                        case "ProductPacksize" when changeEvent.Changes.Any(c => c.PropertyName == "Height" || c.PropertyName == "Width" || c.PropertyName == "Length" || c.PropertyName == "Weight" || c.PropertyName == "BarcodeValue"):
                            ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(ItemDimensionsUpload).FullName));
                            break;
                    }
                }
                catch (Exception e)
                {
                    LogError(e);
                }
            });
            _isPrimed = true;
        }
    }
}