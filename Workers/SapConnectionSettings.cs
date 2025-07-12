using System.Collections.Generic;
using Pro4Soft.SapB1Integration.Infrastructure;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers
{
    public class SapConnectionSettings : Settings
    {
        public List<CompanySettings> Companies { get; set; } = new List<CompanySettings>();
    }

    public class CompanySettings
    {
        public string SLDServer { get; set; }
        public string Server { get; set; }
        public string LicenseServer { get; set; }
        public string CompanyDb { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public bool UseTrusted { get; set; }
        public string Language { get; set; } = BoSuppLangs.ln_English.ToString();
        public string DbServerType { get; set; } = BoDataServerTypes.dst_MSSQL2017.ToString();

        public string ClientName { get; set; }

        public string InvoiceAdjustmentDownloadBin { get; set; }

        public string AdjustmentsPriceList { get; set; }

        public string AdjustmentsPricingSchema { get; set; }

        public string PackagingType { get; set; }

        public Company CreateCompany()
        {
            return new Company
            {
                LicenseServer = LicenseServer,
                SLDServer = SLDServer,
                Server = Server,
                CompanyDB = CompanyDb,
                UserName = UserName,
                Password = Password,
                DbUserName = DbUserName,
                DbPassword = DbPassword,
                UseTrusted = UseTrusted,
                language = Language.ParseEnum<BoSuppLangs>(),
                DbServerType = DbServerType.ParseEnum<BoDataServerTypes>()
            };
        }
    }
}