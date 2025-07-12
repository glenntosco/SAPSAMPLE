using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Dapper;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using Pro4Soft.SapB1Integration.Workers.Upload;
using RestSharp;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers
{
    public abstract class IntegrationBase : BaseWorker
    {
        protected IntegrationBase(ScheduleSetting settings) : base(settings)
        {
        }
        
        //WMS
        private static readonly Lazy<List<Client>> ClientCache = new Lazy<List<Client>>(() =>
        {
            var result = Singleton<EntryPoint>.Instance.WebInvoke("odata/Client", request =>
            {
                request.AddQueryParameter("$select", "Id,Name");
            });
            if (!result.IsSuccessful)
                throw new BusinessWebException(result.StatusCode, result.Content);
            return Utils.DeserializeFromJson<List<Client>>(result.Content, "value");
        });

        private static readonly Dictionary<string, Bin> BinCache = new Dictionary<string, Bin>();
        
        protected Guid? GetClientId(string clientName)
        {
            if (string.IsNullOrWhiteSpace(clientName))
                return null;

            var clientId = ClientCache.Value.SingleOrDefault(c => c.Name == clientName)?.Id;
            if (clientId != null)
                return clientId;

            var client = WebInvoke<List<Client>>("odata/Client", request =>
            {
                request.AddQueryParameter("$filter", $"Name eq '{clientName}'");
                request.AddQueryParameter("$select", "Id,Name");
            }, Method.GET, null, "value").SingleOrDefault();
            if (client == null)
                throw new BusinessWebException($"Client [{clientName}] is not setup");
            ClientCache.Value.Add(client);
            return client.Id;
        }

        protected Bin GetBinDetails(string bin)
        {
            if (string.IsNullOrWhiteSpace(bin))
                return null;

            if (BinCache.TryGetValue(bin, out var descr))
                return descr;

            var binDetls = WebInvoke<List<dynamic>>("odata/Bin", request =>
            {
                request.AddQueryParameter("$filter", $"BinCode eq '{bin}'");
                request.AddQueryParameter("$select", "Id,BinCode");
                request.AddQueryParameter("$expand", "Zone($expand=Warehouse($select=Id,WarehouseCode);$select=Id)");
            }, Method.GET, null, "value").SingleOrDefault();
            if (binDetls == null)
                throw new BusinessWebException($"Bin [{bin}] is not setup");

            var result = new Bin
            {
                BinCode = binDetls.BinCode,
                Id = binDetls.Id,
                WarehouseCode = binDetls.Zone.Warehouse.WarehouseCode
            };

            BinCache[bin] = result;
            return result;
        }

        protected IRestResponse WebInvoke(string url, Action<RestRequest> requestRewrite = null, Method method = Method.GET, dynamic payload = null)
        {
            var result = Singleton<EntryPoint>.Instance.WebInvoke(url, requestRewrite, method, payload);
            if (!result.IsSuccessful)
                throw new BusinessWebException(result.StatusCode, result.Content);
            return result;
        }

        protected T WebInvoke<T>(string url, Action<RestRequest> requestRewrite = null, Method method = Method.GET, object payload = null, string root = null) where T : class
        {
            var resp = Singleton<EntryPoint>.Instance.WebInvoke(url, requestRewrite, method, payload);
            if (!resp.IsSuccessful)
                throw new BusinessWebException(resp.StatusCode, resp.Content);
            return Utils.DeserializeFromJson<T>(resp.Content, root);
        }

        protected Guid? IdLookup(string url, Action<RestRequest> req = null)
        {
            var dataset = WebInvoke<List<dynamic>>(url, request =>
            {
                request.AddQueryParameter("$select", "Id");
                req?.Invoke(request);
            }, Method.GET, null, "value");
            if (!dataset.Any())
                return null;
            if (dataset.Count > 1)
                throw new BusinessWebException($"More than one record found");
            return Guid.TryParse(dataset.SingleOrDefault()?.Id.ToString(), out Guid id) ? id : throw new BusinessWebException($"Cannot parse response");
        }

        //SAP B1
        private Company _currentCompany;
        protected Company CurrentCompany
        {
            get
            {
                if(CompanySettings == null)
                    throw new Exception($"Company settings are not set");
                if (_currentCompany == null)
                    _currentCompany = CompanySettings.CreateCompany();
                if (_currentCompany.Connected)
                    return _currentCompany;
                if (_currentCompany.Connect() == 0)
                    return _currentCompany;

                _currentCompany.GetLastError(out var errorCode, out var errorMessage);
                throw new Exception($"Failed to Connect to SAP B1 Company: {errorCode} - {errorMessage}");
            }
            private set => _currentCompany = value;
        }

        protected CompanySettings CompanySettings { get; private set; }

        protected void SetCompany(CompanySettings settings)
        {
            CompanySettings = settings;
        }

        protected void Disconnect()
        {
            if (_currentCompany?.Connected == true)
                _currentCompany.Disconnect();
            _currentCompany = null;
            CompanySettings = null;
        }

        public double? GetAdjustmentPrice(string itemCode)
        {
            double? result = null;
            switch (CompanySettings.AdjustmentsPricingSchema.ParseEnum<AdjustmentsPricingSchema>(false))
            {
                case AdjustmentsPricingSchema.FixedPriceList:
                    result = ExecuteScalar<double>(
                        $"select price from itm1 where itemcode = '{itemCode}' and pricelist = '{CompanySettings.AdjustmentsPriceList}'");
                    break;
                case AdjustmentsPricingSchema.LastEvaluatedPrice:
                    result = ExecuteScalar<double>(
                        $"select lstevlpric from oitm where itemCode = '{itemCode}'");
                    break;
                case AdjustmentsPricingSchema.LastPurchasePrice:
                    result = ExecuteScalar<double>(
                        $"select lastPurPrc from oitm where itemCode = '{itemCode}'");
                    break;
            }
            return result;
        }

        protected int GetSystemSerial(string serial, string itemCode)
        {
            return ExecuteScalar<int>($@"select SysSerial from osri where IntrSerial = '{serial}' and ItemCode = '{itemCode}' and Status = 0");
        }

        protected double GetNumPerMsr(int docEntry, int lineNum, string tableName)
        {
            return ExecuteScalar<double>($@"
select
   numpermsr 
from 
    {tableName}
where
	docentry = {docEntry} and linenum = {lineNum}");
        }

        protected T ExecuteScalar<T>(string sql)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                return db.ExecuteScalar<T>(sql);
        }

        protected List<dynamic> ExecuteReader(string sql)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                return db.Query(sql).ToList();
        }

        //Staging table
        public void StagingTableLockUnlock(string id, int objType, bool lockObj)
        {
            if (string.IsNullOrWhiteSpace(id))
                return;

            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                db.Execute(@"
update 
    [@P4S_EXPORT]
set
    U_Lock = @Lock,
    U_DateModified = @DateModified
where
    U_ObjType = @ObjType and
    U_DocId = @DocId", new
                {
                    Lock = lockObj ? 1 : 0,
                    ObjType = objType,
                    DocId = id,
                    DateModified = DateTime.Now
                });
        }

        public void StagingTableDelete(string id, int objType)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                db.Execute(@"
delete from 
    [@P4S_EXPORT]
where
    U_ObjType = @ObjType and
    U_DocId = @DocId", new
                {
                    ObjType = objType,
                    DocId = id,
                });
        }

        public void StagingTableIncrementBackOrder(string id, int objType)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                db.Execute(@"
update 
    [@P4S_EXPORT]
set
    U_BackOrder = U_BackOrder+1,
    U_DateModified = @DateModified,
    U_ExportFlag = 0
where
    U_ObjType = @ObjType and
    U_DocId = @DocId", new
                {
                    ObjType = objType,
                    DocId = id,
                    DateModified = DateTime.Now
                });
        }

        public bool IsWoReceiptUploaded(int woKey)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                return db.ExecuteScalar<int>(@"
select 
    U_ExportFlag 
from 
    [@P4S_EXPORT] 
where 
    U_ObjType = 202 and
    U_DocId = @DocId", new
                {
                    DocId = woKey,
                }) == 3;
        }

        public void StagingSetWoReceiptUploaded(int woKey)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                db.Execute(@"
update 
    [@P4S_EXPORT]
set
    U_DateModified = @DateModified,
    U_ExportFlag = @ExportFlag
where
    U_ObjType = 202 and
    U_DocId = @DocId", new
                {
                    DocId = woKey,
                    DateModified = DateTime.Now,
                    ExportFlag = 3
                });
        }

        public void StagingTableMarkDownloaded(string id, int objType, string errorMessage = null)
        {
            using (IDbConnection db = new SqlConnection($"Server={CompanySettings.Server};Initial Catalog={CompanySettings.CompanyDb};Persist Security Info=False;User ID={CompanySettings.DbUserName};Password={CompanySettings.DbPassword};"))
                db.Execute(@"
update 
    [@P4S_EXPORT]
set
    U_DateModified = @DateModified,
    U_ExportFlag = @ExportFlag,
    U_Error = @ErrorMessage
where
    U_ObjType = @ObjType and
    U_DocId = @DocId", new
                {
                    ObjType = objType,
                    DocId = id,
                    DateModified = DateTime.Now,
                    ExportFlag = string.IsNullOrWhiteSpace(errorMessage) ? 1 : 2,
                    ErrorMessage = string.IsNullOrWhiteSpace(errorMessage) ? null : errorMessage,
                });
        }
    }


}