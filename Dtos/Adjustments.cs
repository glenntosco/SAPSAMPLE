using System;
using System.Collections.Generic;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class Adjustment
    {
        public Guid Id { get; set; }

        public DateTimeOffset Timestamp { get; set; }
        public string Client { get; set; }
        public string SubType { get; set; }

        public string FromWarehouse { get; set; }
        public string ToWarehouse { get; set; }
        public string Sku { get; set; }

        public int? EachCount { get; set; }
        public int? NumberOfPacks { get; set; }
        public string LotNumber { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public string SerialNumber { get; set; }

        public decimal Quantity { get; set; }
        public string Reason { get; set; }
    }

    public class ReturnAdjustment
    {
        public Guid Id { get; set; }

        public DateTimeOffset Timestamp { get; set; }

        public string Client { get; set; }

        public string ReferenceType { get; set; } = null;
        public Guid? ReferenceId { get; set; } = null;

        public string FromWarehouse { get; set; }
        public string ToWarehouse { get; set; }
        public string Sku { get; set; }

        public int? EachCount { get; set; }
        public int? NumberOfPacks { get; set; }
        public string LotNumber { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public string SerialNumber { get; set; }

        public decimal Quantity { get; set; }
        public string Reason { get; set; }
    }

    public class InventoryHelper
    {
        public string Warehouse { get; set; }
        public string Sku { get; set; }
        public decimal Quantity { get; set; }
        public bool IsPacksizeControlled { get; set; }

        public List<InventoryDetailsHelper> Details { get; set; } = new List<InventoryDetailsHelper>();
    }

    public class InventoryDetailsHelper
    {
        public string SerialNumber { get; set; }
        public string LotNumber { get; set; }
        public DateTimeOffset? Expiry { get; set; }
        public int? EachCount { get; set; }
        public decimal Quantity { get; set; }
    }
}
