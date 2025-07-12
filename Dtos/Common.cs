using System;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class IdObject
    {
        public Guid Id { get; set; }
    }

    public class Client: IdObject
    {
        public string Name { get; set; }
    }

    public class Customer: IdObject
    {
        public string CustomerCode { get; set; }
        public string Email { get; set; }
    }

    public class VendorSimpleObject: IdObject
    {
        public string VendorCode { get; set; }
    }
    
    public class Bin : IdObject
    {
        public string BinCode { get; set; }
        public string WarehouseCode { get; set; }
    }
}