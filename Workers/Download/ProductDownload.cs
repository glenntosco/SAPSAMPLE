using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;


namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public abstract class ProductDownload: IntegrationBase
    {
        protected ProductDownload(ScheduleSetting settings) : base(settings)
        {
        }

        protected List<FullProduct> MapFromB1(Guid? clientId, List<string> itemCodes)
        {
            var result = ExecuteReader($@"
select
    OITM.ItemCode,
    OITM.ItemName, 
    OITM.CodeBars,
    OITB.ItmsGrpNam,
    OBCD.BcdCode,
    OITM.ManSerNum, 
    OITM.ManBtchNum, 
    OITM.U_P4S_IsExpiry, 
    OITM.U_P4S_IsDecimal,
    OITM.BLength1, 
    OITM.BHeight1, 
    OITM.BWeight1, 
    OITM.BWidth1, 
    OUOM.UomName,
    UGP1.BaseQty/UGP1.AltQty as Packsize,
    ITM12.Length1, 
    ITM12.Width1, 
    ITM12.Height1,
	ITM12.Weight1    
from 
    OITM	
    left outer join UGP1 on UGP1.UgpEntry = OITM.UgpEntry
    left outer join OITB on OITB.ItmsGrpCod = OITM.ItmsGrpCod
    left outer join OUOM on UGP1.UomEntry = OUOM.UomEntry
    left outer join ITM12 on ITM12.ItemCode = OITM.ItemCode and OUOM.UomEntry = ITM12.UomEntry and ITM12.UomType = 'P'
    left outer join OBCD on OBCD.ItemCode = OITM.ItemCode and OUOM.UomEntry = OBCD.UomEntry
where 
    --ITM12.UoMType = 'P' and
    OITM.InvntItem = 'Y' and
    OITM.Canceled = 'N' and
    OITM.ItemCode in ({string.Join(",", itemCodes.Select(c => $"'{c}'"))})
order by
    Packsize")
                .GroupBy(c => c.ItemCode)
                .Select(c => new
                {
                    item = c.FirstOrDefault(),
                    group = c
                })
                .Select(c => new FullProduct
                {
                    Sku = c.item.ItemCode,
                    Category = c.item.ItmsGrpNam,
                    Description = c.item.ItemName,
                    Upc = c.item.CodeBars,
                    Length = c.item.Length1 ?? c.item.BLength1,
                    Width = c.item.Width1 ?? c.item.BWidth1,
                    Height = c.item.Height1 ?? c.item.BHeight1,
                    Weight = c.item.Weight1 ?? c.item.BWeight1,
                    ClientId = clientId,
                    IsLotControlled = c.item.ManBtchNum == "Y",
                    IsSerialControlled = c.item.ManSerNum == "Y",
                    IsExpiryControlled = c.item.U_P4S_IsExpiry == 1,
                    IsDecimalControlled = c.item.U_P4S_IsDecimal == 1,
                    IsPacksizeControlled = c.group.Any(c1 => c1.Packsize > 1),
                    Packsizes = c.group
                        .Select(c1 => new ProductPacksize
                        {
                            EachCount = (int) c1.Packsize,
                            Name = c1.UomName,
                            Description = c.item.ItemName,
                            BarcodeValue = c1.BcdCode,
                            Length = (decimal?)c1.Length1,
                            Width = (decimal?)c1.Width1,
                            Height = (decimal?)c1.Height1,
                            Weight = (decimal?)c1.Weight1,
                        })
                        .ToList()
                }).ToList();

            result.ForEach(product =>
            {
                product.Components = ExecuteReader($@"
select 
	OITT.Code as Sku, -- Parent
	ITT1.Code as ComponentSku, -- Child
	ITT1.Quantity/OITT.Qauntity as ComponentQuantity  -- Child
from 
	OITT 
	left outer join ITT1 on ITT1.Father = OITT.Code
where 
	OITT.TreeType = 'P' and
	ITT1.IssueMthd = 'M' and
	OITT.Code = '{product.Sku}'").Select(c => new ProductComponent
                {
                    ComponentProduct = new Product {Sku = c.ComponentSku},
                    Quantity = c.ComponentQuantity
                }).ToList();
                product.IsBillOfMaterial = product.Components.Any();
            });

            return result;
        }

        protected FullProduct CreateUpdateProduct(FullProduct b1Prod, Guid? existingId)
        {
            //if (existing == null)
            //    return null;
            dynamic payload = new ExpandoObject();
            payload.Id = existingId;
            payload.ClientId = b1Prod.ClientId;
            payload.Sku = b1Prod.Sku;
            payload.Category = b1Prod.Category;
            payload.Upc = b1Prod.Upc;
            payload.Description = b1Prod.Description;
            payload.IsBillOfMaterial = b1Prod.IsBillOfMaterial;
            payload.IsSerialControlled = b1Prod.IsSerialControlled;
            payload.IsLotControlled = b1Prod.IsLotControlled;
            payload.IsPacksizeControlled = b1Prod.IsPacksizeControlled;
            payload.IsExpiryControlled = b1Prod.IsExpiryControlled;
            payload.IsDecimalControlled = b1Prod.IsDecimalControlled;

            payload.Length = b1Prod.Length;
            payload.Width = b1Prod.Width;
            payload.Height = b1Prod.Height;
            payload.Weight = b1Prod.Weight;
            if (b1Prod.IsPacksizeControlled)
            {
                var eachPacksize = b1Prod.Packsizes.SingleOrDefault(c => c.EachCount == 1);
                payload.Upc = eachPacksize?.BarcodeValue;
                payload.Length = eachPacksize?.Length;
                payload.Width = eachPacksize?.Width;
                payload.Height = eachPacksize?.Height;
                payload.Weight = eachPacksize?.Weight;
            }

            FullProduct createdUpdatedProduct = WebInvoke<FullProduct>("api/ProductApi/CreateOrUpdate", null, Method.POST, payload);
            Log($"Sku: [{b1Prod.Sku}] " + (existingId == null ? "create" : "updated"));
            if (b1Prod.IsPacksizeControlled)
            {
                foreach (var packsize in b1Prod.Packsizes.Where(c => c.EachCount != 1))
                {
                    var existingPacksize = createdUpdatedProduct.Packsizes.SingleOrDefault(c => c.Name == packsize.Name);
                    WebInvoke("api/ProductApi/AddOrUpdatePacksize", req =>
                    {
                        req.AddQueryParameter("productId", createdUpdatedProduct.Id.ToString());
                    }, Method.POST, new
                    {
                        existingPacksize?.Id,
                        packsize.BarcodeValue,
                        packsize.Name,
                        packsize.Description,
                        packsize.EachCount,
                        packsize.Length,
                        packsize.Width,
                        packsize.Height,
                        packsize.Weight,
                    });

                    Log($"Packsize: [{packsize.Name}] created/updated");
                    if (existingPacksize != null)
                        createdUpdatedProduct.Packsizes.Remove(existingPacksize);
                }

                if (createdUpdatedProduct.Packsizes.Any())
                    foreach (var oldPacksizesin in createdUpdatedProduct.Packsizes.Where(c => c.EachCount != 1))
                    {
                        WebInvoke("api/ProductApi/RemovePacksize", req => { req.AddQueryParameter("packsizeId", oldPacksizesin.Id.ToString()); });
                        Log($"Packsize: [{oldPacksizesin.Name}] removed");
                    }
            }

            if (b1Prod.IsBillOfMaterial)
            {
                foreach (var component in b1Prod.Components)
                {
                    var existingComponent = createdUpdatedProduct.Components.SingleOrDefault(c => c.ComponentProduct.Sku == component.ComponentProduct.Sku);
                    WebInvoke("api/ProductApi/AddOrUpdateComponentBySku", req =>
                    {
                        req.AddQueryParameter("productId", createdUpdatedProduct.Id.ToString());
                        req.AddQueryParameter("componentSku", component.ComponentProduct.Sku);
                        req.AddQueryParameter("quantity", component.Quantity.ToString());
                    });

                    Log($"BOM Component: [{component.ComponentProduct.Sku}] created/updated");
                    if (existingComponent != null)
                        createdUpdatedProduct.Components.Remove(existingComponent);
                }

                if (createdUpdatedProduct.Components.Any())
                    foreach (var oldComponent in createdUpdatedProduct.Components)
                    {
                        WebInvoke("api/ProductApi/RemoveComponent", req => { req.AddQueryParameter("componentId", oldComponent.Id.ToString()); });
                        Log($"Component: [{oldComponent.ComponentProduct.Sku}] removed");
                    }
            }

            return createdUpdatedProduct;
        }
    }
}