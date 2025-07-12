using System;
using System.Collections.Generic;

namespace Pro4Soft.SapB1Integration.Dtos
{
    public class FullProduct: IdObject
    {
        public string Sku { get; set; }
        public string Upc { get; set; }
        public string Category { get; set; }
        public string Description { get; set; }

        public string ReferenceNumber { get; set; }

        public bool IsBillOfMaterial { get; set; }
        public bool IsLotControlled { get; set; }
        public bool IsSerialControlled { get; set; }
        public bool IsExpiryControlled { get; set; }
        public bool IsDecimalControlled { get; set; }
        public bool IsPacksizeControlled { get; set; }

        public decimal? Height { get; set; }
        public decimal? Width { get; set; }
        public decimal? Length { get; set; }
        public decimal? Weight { get; set; }

        public Client Client { get; set; }
        public Guid? ClientId { get; set; }

        public string DecimalTrackType { get; set; }
        public string WeightUnitOfMeasure { get; set; }
        public string LengthUnitOfMeasure { get; set; }

        public string DimsLengthUnitOfMeasure { get; set; }
        public string DimsWeightUnitOfMeasure { get; set; }

        public List<ProductComponent> Components { get; set; } = new List<ProductComponent>();
        public List<ProductPacksize> Packsizes { get; set; } = new List<ProductPacksize>();
    }

    public class ProductComponent : IdObject
    {
        public decimal Quantity { get; set; }
        public Product ComponentProduct { get; set; }
    }

    public class ProductPacksize : IdObject
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string BarcodeValue { get; set; }

        public FullProduct Product { get; set; }

        public int EachCount { get; set; }

        public decimal? Height { get; set; }
        public decimal? Width { get; set; }
        public decimal? Length { get; set; }
        public decimal? Weight { get; set; }
    }

    public class Product : IdObject
    {
        public string Sku { get; set; }
    }
}
