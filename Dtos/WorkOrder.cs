using System;
using System.Collections.Generic;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class WorkOrder : IdObject
    {
        public string WorkOrderNumber { get; set; }
        public string ReferenceNumber { get; set; }
        public string WorkOrderState { get; set; }
        public decimal ProducedQuantity { get; set; }
        public Client Client { get; set; }
        public Guid? ClientId { get; set; }
        public Product Product { get; set; }
        public List<WorkOrderFinishedProductDetail> ProductDetails { get; set; } = new List<WorkOrderFinishedProductDetail>();
        public List<WorkOrderComponentLine> Lines { get; set; } = new List<WorkOrderComponentLine>();
        public DateTimeOffset? UploadDate { get; set; }
    }

    public class WorkOrderComponentLine : IdObject
    {
        public int LineNumber { get; set; }
        public string ReferenceNumber { get; set; }
        public Product Product { get; set; }
        public decimal ConsumedQuantity { get; set; }
        public List<WorkOrderComponentLineDetail> ConsumedProductDetails { get; set; } = new List<WorkOrderComponentLineDetail>();
    }

    public class WorkOrderComponentLineDetail : IdObject
    {
        public int? PacksizeEachCount { get; set; }
        public int? ConsumedQuantity { get; set; }
        public string LotNumber { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public string SerialNumber { get; set; }
    }

    public class WorkOrderFinishedProductDetail : IdObject
    {
        public int? PacksizeEachCount { get; set; }
        public int? ProducedQuantity { get; set; }
        public string LotNumber { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public string SerialNumber { get; set; }
    }

    
}
