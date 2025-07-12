using System;
using System.Collections.Generic;
using System.Linq;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class SalesOrder : IdObject
    {
        public string PickTicketNumber { get; set; }
        public string ReferenceNumber { get; set; }
        public string PickTicketState { get; set; }
        public Client Client { get; set; }
        public Guid? ClientId { get; set; }
        public List<SalesOrderTote> Totes { get; set; } = new List<SalesOrderTote>();
        public DateTimeOffset? UploadDate { get; set; }
        public bool IsWarehouseTransfer { get; set; }

        public List<SalesOrderToteLine> GetOrderLines()
        {
            var result = new List<SalesOrderToteLine>();
            foreach (var salesOrderTote in Totes)
            {
                foreach (var toteLine in salesOrderTote.Lines)
                {
                    var existing = result.SingleOrDefault(c => c.PickTicketLine.ReferenceNumber == toteLine.PickTicketLine.ReferenceNumber);
                    if (existing == null)
                    {
                        existing = new SalesOrderToteLine {PickTicketLine = toteLine.PickTicketLine};
                        result.Add(existing);
                    }

                    existing.Product = toteLine.Product;
                    existing.PickedQuantity += toteLine.PickedQuantity;
                    existing.LineDetails.AddRange(toteLine.LineDetails);
                }
            }
            return result;
        }
    }

    public class SalesOrderTote : IdObject
    {
        public string Sscc18Code { get; set; }
        public int CartonNumber { get; set; }
        public List<SalesOrderToteLine> Lines { get; set; } = new List<SalesOrderToteLine>();
    }

    public class SalesOrderToteLine : IdObject
    {
        public decimal PickedQuantity { get; set; }
        public SalesOrderLine PickTicketLine { get; set; }
        public Product Product { get; set; }
        public List<SalesOrderToteLineDetails> LineDetails { get; set; } = new List<SalesOrderToteLineDetails>();
    }

    public class SalesOrderLine : IdObject
    {
        public int LineNumber { get; set; }
        public string ReferenceNumber { get; set; }
    }

    public class SalesOrderToteLineDetails : IdObject
    {
        public int? PacksizeEachCount { get; set; }
        public string LotNumber { get; set; }
        public string SerialNumber { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public decimal PickedQuantity { get; set; }
    }
}
