using System;
using System.Runtime.InteropServices;
using Pro4Soft.SapB1Integration.Infrastructure;
using SAPbobsCOM;

namespace Pro4Soft.SapB1Integration.Workers
{
    public class Install : IntegrationBase
    {
        public Install(ScheduleSetting settings) : base(settings)
        {
        }

        public override void Execute()
        {
            foreach (var companySettings in App<SapConnectionSettings>.Instance.Companies)
            {
                SetCompany(companySettings);
                var userTable = (UserTablesMD)CurrentCompany.GetBusinessObject(BoObjectTypes.oUserTables);
                try
                {
                    if (!userTable.GetByKey("P4S_Export"))
                    {
                        userTable.TableName = "P4S_Export";
                        userTable.TableType = BoUTBTableType.bott_NoObject;
                        userTable.TableDescription = "Pro4Soft Export table";

                        if (userTable.Add() != 0)
                        {
                            CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                            throw new Exception($"{errorCode} - {errorMessage}");
                        }

                        Marshal.ReleaseComObject(userTable);

                        AddFieldToExportTable("DateCreated", BoFieldTypes.db_Date, null);
                        AddFieldToExportTable("DateModified", BoFieldTypes.db_Date, null);
                        AddFieldToExportTable("DocId", BoFieldTypes.db_Alpha, 20);
                        AddFieldToExportTable("ObjType", BoFieldTypes.db_Numeric, 11);
                        AddFieldToExportTable("ExportFlag", BoFieldTypes.db_Numeric, 11);
                        AddFieldToExportTable("Lock", BoFieldTypes.db_Numeric, 11);
                        AddFieldToExportTable("Error", BoFieldTypes.db_Alpha, 254);
                        AddFieldToExportTable("BackOrder", BoFieldTypes.db_Numeric, 11);
                    }
                    else
                    {
                        Log($"Staging table already Installed for company {companySettings.ClientName ?? companySettings.CompanyDb}");
                    }

                    if (ExecuteScalar<int>("select count(*) from CUFD where tableID = 'OITM' and AliasID = 'P4S_IsExpiry'") == 0)
                    {
                        Marshal.ReleaseComObject(userTable);
                        AddFieldToOitmTable("P4S_IsExpiry", "Expiry Date Tracking");
                    }
                    else
                    {
                        Log($"P4S_IsExpiry field already Installed for company {companySettings.ClientName ?? companySettings.CompanyDb}");
                    }

                    if (ExecuteScalar<int>("select count(*) from CUFD where tableID = 'OITM' and AliasID = 'P4S_IsDecimal'") == 0)
                    {
                        Marshal.ReleaseComObject(userTable);
                        AddFieldToOitmTable("P4S_IsDecimal", "Decimal Tracking");
                    }
                    else
                    {
                        Log($"P4S_IsDecimal field already Installed for company {companySettings.ClientName ?? companySettings.CompanyDb}");
                    }
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
                finally
                {
                    if (userTable != null)
                        Marshal.ReleaseComObject(userTable);
                    Disconnect();
                }
            }
        }

        private void AddFieldToExportTable(string fieldName, BoFieldTypes fieldType, int? fieldSize)
        {
            var userField = (UserFieldsMD)CurrentCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            try
            {
                userField.TableName = "P4S_Export";
                userField.Name = fieldName;
                userField.Type = fieldType;
                userField.EditSize = fieldSize ?? userField.EditSize;
                if (userField.Add() == 0)
                    return;

                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                throw new Exception($"SAP B1: {errorCode} - {errorMessage}");
            }
            finally
            {
                if (userField != null)
                    Marshal.ReleaseComObject(userField);
            }
        }

        private void AddFieldToOitmTable(string name, string description)
        {
            var userField = (UserFieldsMD) CurrentCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            try
            {
                userField.TableName = "OITM";
                userField.Name = name;
                userField.Description = description;
                userField.Type = BoFieldTypes.db_Numeric;
                userField.EditSize = 10;

                userField.DefaultValue = "0";
                userField.ValidValues.SetCurrentLine(0);
                userField.ValidValues.Value = "0";
                userField.ValidValues.Description = "Not Tracked";
                userField.ValidValues.Add();
                userField.ValidValues.SetCurrentLine(1);
                userField.ValidValues.Value = "1";
                userField.ValidValues.Description = "Tracked";

                if (userField.Add() == 0)
                    return;

                CurrentCompany.GetLastError(out var errorCode, out var errorMessage);
                throw new Exception($"SAP B1: {errorCode} - {errorMessage}");
            }
            finally
            {
                if (userField != null)
                    Marshal.ReleaseComObject(userField);
            }
        }
    }
}