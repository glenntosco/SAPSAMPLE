using System.Collections.Generic;
using System;
using System.Linq;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
// ReSharper disable AssignNullToNotNullAttribute
// ReSharper disable PossibleNullReferenceException

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class ProductFullDownload : ProductDownload
    {
        public ProductFullDownload(ScheduleSetting settings) : base(settings)
        {

        }

        public override void Execute()
        {
            var existingProds = WebInvoke<List<FullProduct>>("odata/Product", request =>
            {
                request.AddQueryParameter("$select", "Id,Sku,ClientId");
            }, Method.GET, null, "value");

            var b1Products = new List<FullProduct>();

            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                try
                {
                    SetCompany(companySettings);
                    var data = ExecuteReader(@"
select
	ItemCode
from 
    OITM
where
	InvntItem = 'Y'");
                    if (!data.Any())
                        continue;
                    b1Products.AddRange(MapFromB1(GetClientId(companySettings.ClientName), data.Select(c => c.ItemCode).Cast<string>().ToList()));
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

            foreach (var b1Prod in b1Products.ToList())
            {
                try
                {
                    var existing = existingProds.Where(c => c.ClientId == b1Prod.ClientId).SingleOrDefault(c => c.Sku == b1Prod.Sku);
                    CreateUpdateProduct(b1Prod, existing?.Id);
                    existingProds.Remove(existing);
                }
                catch (Exception e)
                {
                    LogError(e);
                }
            }

            if (existingProds.Any(c => c != null))
            {
                WebInvoke("api/ProductApi/DeleteBatch", null, Method.POST, b1Products.Where(c => c != null).Select(c => c.Id).ToList());
                Log($"[{b1Products.Count}] products removed");
            }
        }
    }
}