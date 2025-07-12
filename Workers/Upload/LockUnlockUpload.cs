using System;
using System.Collections.Generic;
using System.Linq;
using Pro4Soft.SapB1Integration.Dtos;
using Pro4Soft.SapB1Integration.Infrastructure;
using RestSharp;
using SAPbobsCOM;

// ReSharper disable InconsistentNaming

namespace Pro4Soft.SapB1Integration.Workers.Upload
{
    public class LockUnlockUpload : IntegrationBase
    {
        private static bool _isPrimed;
        private const string StateChangeUploadLastTime = nameof(StateChangeUploadLastTime);

        public LockUnlockUpload(ScheduleSetting settings) : base(settings)
        {

        }

        public override void Execute()
        {
            Subscribe();

            var now = DateTimeOffset.UtcNow;
            var serverVal = WebInvoke<string>("api/TenantApi/GetTenantSettings", request => { request.AddQueryParameter("key", StateChangeUploadLastTime); });
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
                                Value = "PickTicket",
                                Operator = "Equal"
                            },
                            new
                            {
                                Field = "ObjectType",
                                Value = "PurchaseOrder",
                                Operator = "Equal"
                            },
                            new
                            {
                                Field = "ObjectType",
                                Value = "WorkOrder",
                                Operator = "Equal"
                            },
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

            var whTransferUpload = false;
            var poUpload = false;
            var soUpload = false;
            var woUpload = false;
            var interestedFields = new[] {"PickTicketState", "PurchaseOrderState", "WorkOrderState", "UploadDate"};
            foreach (var objectTypeGroup in historyRecords
                .Where(c => c.Changes.Any(c1 => interestedFields.Contains(c1.PropertyName)))
                .GroupBy(c => c.ObjectType))
            {
                var distinctObjs = objectTypeGroup.Select(c => c.ObjectId).Distinct().ToList();

                switch (objectTypeGroup.Key)
                {
                    case "PickTicket":
                    {
                        var sos = new List<SalesOrder>();
                        for (var i = 0; i < distinctObjs.Count; i += 10)
                        {
                            var batch = distinctObjs.Skip(i).Take(10);
                            sos.AddRange(WebInvoke<List<SalesOrder>>("odata/PickTicket", request =>
                            {
                                request.AddQueryParameter("$select", @"Id,IsWarehouseTransfer,PickTicketNumber,ReferenceNumber,PickTicketState,UploadDate");
                                request.AddQueryParameter("$expand", @"Client($select=Name)");
                                request.AddQueryParameter("$filter", string.Join(" or ", batch.Select(c => $"(Id eq {c})")));
                            }, Method.GET, null, "value"));
                        }

                        foreach (var clientSos in sos.GroupBy(c => c.Client?.Name))
                        {
                            var companySettings = App<SapConnectionSettings>.Instance.Companies.SingleOrDefault(c => c.ClientName == clientSos.Key);
                            if (companySettings == null)
                                continue;
                            SetCompany(companySettings);
                            foreach (var so in clientSos)
                            {
                                try
                                {
                                    switch (so.PickTicketState)
                                    {
                                        case "Draft":
                                        case "Unallocated":
                                        case "Closed":
                                            if (!so.IsWarehouseTransfer)
                                            {
                                                StagingTableLockUnlock(so.ReferenceNumber, (int) BoObjectTypes.oOrders, false);
                                                Log($"SO: [{so.PickTicketNumber}] unlocked");
                                            }
                                            else if(so.PickTicketState == "Unallocated")
                                            {
                                                StagingTableLockUnlock(so.ReferenceNumber, (int)BoObjectTypes.oInventoryTransferRequest, false);
                                                Log($"WH Transfer: [{so.PickTicketNumber}] unlocked");
                                            }
                                            break;
                                        default:
                                            if (!so.IsWarehouseTransfer)
                                            {
                                                StagingTableLockUnlock(so.ReferenceNumber, (int) BoObjectTypes.oOrders, true);
                                                Log($"SO: [{so.PickTicketNumber}] locked");
                                            }
                                            else
                                            {
                                                StagingTableLockUnlock(so.ReferenceNumber, (int)BoObjectTypes.oInventoryTransferRequest, true);
                                                Log($"WH Transfer [{so.PickTicketNumber}] locked");
                                            }
                                            break;
                                    }
                                    if (!so.IsWarehouseTransfer)
                                        soUpload = soUpload || so.PickTicketState == "Closed" && so.UploadDate == null;
                                }
                                catch (Exception e)
                                {
                                    LogError(e);
                                }
                            }

                            Disconnect();
                        }

                        break;
                    }
                    case "PurchaseOrder":
                    {
                        var pos = new List<PurchaseOrder>();
                        for (var i = 0; i < distinctObjs.Count; i += 10)
                        {
                            var batch = distinctObjs.Skip(i).Take(10);
                            pos.AddRange(WebInvoke<List<PurchaseOrder>>("odata/PurchaseOrder", request =>
                            {
                                request.AddQueryParameter("$select", @"Id,IsWarehouseTransfer,PurchaseOrderNumber,ReferenceNumber,PurchaseOrderState,UploadDate");
                                request.AddQueryParameter("$expand", @"Client($select=Name)");
                                request.AddQueryParameter("$filter", string.Join(" or ", batch.Select(c => $"(Id eq {c})")));
                            }, Method.GET, null, "value"));
                        }

                        foreach (var clientPos in pos.GroupBy(c => c.Client?.Name))
                        {
                            var companySettings = App<SapConnectionSettings>.Instance.Companies.SingleOrDefault(c => c.ClientName == clientPos.Key);
                            if (companySettings == null)
                                continue;
                            SetCompany(companySettings);
                            foreach (var po in clientPos)
                            {
                                try
                                {
                                    switch (po.PurchaseOrderState)
                                    {
                                        case "Draft":
                                        case "NotReceived":
                                        case "Closed":
                                            if (!po.IsWarehouseTransfer)
                                            {
                                                StagingTableLockUnlock(po.ReferenceNumber, (int) BoObjectTypes.oPurchaseOrders, false);
                                                Log($"PO [{po.PurchaseOrderNumber}] unlocked");
                                            }
                                            break;
                                        default:
                                            if (!po.IsWarehouseTransfer)
                                            {
                                                StagingTableLockUnlock(po.ReferenceNumber, (int) BoObjectTypes.oPurchaseOrders, true);
                                                Log($"PO [{po.PurchaseOrderNumber}] locked");
                                            }
                                            break;
                                    }

                                    if (!po.IsWarehouseTransfer)
                                        poUpload = poUpload || po.PurchaseOrderState == "Closed" && po.UploadDate == null;
                                    else
                                        whTransferUpload = whTransferUpload || po.PurchaseOrderState == "Closed" && po.UploadDate == null;
                                }
                                catch (Exception e)
                                {
                                    LogError(e);
                                }
                            }

                            Disconnect();
                        }

                        break;
                    }
                    case "WorkOrder":
                    {
                        var wos = new List<WorkOrder>();
                        for (var i = 0; i < distinctObjs.Count; i += 10)
                        {
                            var batch = distinctObjs.Skip(i).Take(10);
                            wos.AddRange(WebInvoke<List<WorkOrder>>("odata/WorkOrder", request =>
                            {
                                request.AddQueryParameter("$select", @"Id,WorkOrderNumber,ReferenceNumber,WorkOrderState,UploadDate");
                                request.AddQueryParameter("$expand", @"Client($select=Name)");
                                request.AddQueryParameter("$filter", string.Join(" or ", batch.Select(c => $"(Id eq {c})")));
                            }, Method.GET, null, "value"));
                        }

                        foreach (var clientWos in wos.GroupBy(c => c.Client?.Name))
                        {
                            var companySettings = App<SapConnectionSettings>.Instance.Companies.SingleOrDefault(c => c.ClientName == clientWos.Key);
                            if (companySettings == null)
                                continue;
                            SetCompany(companySettings);
                            foreach (var wo in clientWos)
                            {
                                try
                                {
                                    switch (wo.WorkOrderState)
                                    {
                                        case "Draft":
                                        case "Unallocated":
                                        case "Closed":
                                            StagingTableLockUnlock(wo.ReferenceNumber, (int)BoObjectTypes.oProductionOrders, false);
                                            Log($"WO [{wo.WorkOrderNumber}] unlocked");
                                            break;
                                        default:
                                            StagingTableLockUnlock(wo.ReferenceNumber, (int)BoObjectTypes.oProductionOrders, true);
                                            Log($"PO [{wo.WorkOrderNumber}] locked");
                                            break;
                                    }

                                    woUpload = woUpload || wo.WorkOrderState == "Closed" && wo.UploadDate == null;
                                }
                                catch (Exception e)
                                {
                                    LogError(e);
                                }
                            }

                            Disconnect();
                        }

                        break;
                    }
                }
            }

            WebInvoke("api/TenantApi/SetTenantSettings", request => { request.AddQueryParameter("key", StateChangeUploadLastTime); }, Method.POST, now.ToString());

            if (poUpload)
                ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(PurchaseOrderUpload).FullName));
            if (whTransferUpload)
                ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(WarehouseTransferUpload).FullName));
            if (soUpload)
                ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(SalesOrderUpload).FullName));
            if (woUpload)
                ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(WorkOrderUpload).FullName));
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
                        case "WorkOrder" when changeEvent.Changes.Any(c => c.PropertyName == "WorkOrderState" || c.PropertyName == "UploadDate"): 
                        case "PickTicket" when changeEvent.Changes.Any(c => c.PropertyName == "PickTicketState" || c.PropertyName == "UploadDate"):
                        case "PurchaseOrder" when changeEvent.Changes.Any(c => c.PropertyName == "PurchaseOrderState" || c.PropertyName == "UploadDate"):
                            ScheduleThread.Instance.RunTask(App<Settings>.Instance.Schedules.SingleOrDefault(c => c.Class == typeof(LockUnlockUpload).FullName));
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