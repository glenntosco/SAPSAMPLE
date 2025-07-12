using System.Collections.Generic;
using System;
using System.Linq;
using Newtonsoft.Json;
using SAPbobsCOM;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;

namespace Pro4Soft.SapB1Integration.Workers.Download
{
    public class ProductDeltaDownload : ProductDownload
    {
        public ProductDeltaDownload(ScheduleSetting settings) : base(settings)
        {
            
        }

        public override void Execute()
        {
            var b1Products = new List<FullProduct>();

            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                SetCompany(companySettings);
                try
                {
                    var docEntries = new List<string>();
                    var recordset = ExecuteReader($@"
select
	U_DocId
from [@P4S_EXPORT]
where
	U_ExportFlag = 0 and U_ObjType = {(int) BoObjectTypes.oItems}");
                    if (!recordset.Any())
                        continue;
                    docEntries.AddRange(recordset.Select(c=>c.U_DocId?.ToString()).Cast<string>());
                    b1Products.AddRange(MapFromB1(GetClientId(companySettings.ClientName), docEntries));

                    foreach (var product in b1Products)
                    {
                        try
                        {
                            var existingProd = WebInvoke<List<FullProduct>>("odata/Product", request =>
                            {
                                if (string.IsNullOrWhiteSpace(companySettings.ClientName))
                                    request.AddQueryParameter("$filter", $"Sku eq '{product.Sku}' and ClientId eq null");
                                else
                                    request.AddQueryParameter("$filter", $"Sku eq '{product.Sku}' and Client/Name eq '{companySettings.ClientName}'");
                                request.AddQueryParameter("$select", "Id,ClientId");
                                request.AddQueryParameter("$expand", "Components($expand=ComponentProduct),Packsizes");
                            }, Method.GET, null, "value").SingleOrDefault();

                            CreateUpdateProduct(product, existingProd?.Id);

                            StagingTableDelete(product.Sku, (int) BoObjectTypes.oItems);
                        }
                        catch (Exception e)
                        {
                            StagingTableMarkDownloaded(product.Sku, (int)BoObjectTypes.oItems, e.ToString().Trim().Truncate(254));
                            LogError($@"Item [{product.Sku}] failed to download: {e}
----------
{Utils.SerializeToStringJson(product, Formatting.Indented)}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Items failed to load: {ex}");
                }
                finally
                {
                    Disconnect();
                }
            }
        }
    }
}