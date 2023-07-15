namespace ExportConsoleApp.Models
{
    public class InvoiceRedBubbleModel
    {
        public List<InvoiceRedBubbleForGarmentModel> Garments = new List<InvoiceRedBubbleForGarmentModel>();
        public List<InvoiceRedBubbleForPackagingModel> Packaging = new List<InvoiceRedBubbleForPackagingModel>();

    }

    public class InvoiceRedBubbleForGarmentModel
    {
        public string? InvoiceDate { get; set; }
        public string? InvoiceNumber { get; set; }
        public string? FulfillerName { get; set; }
        public string? Facility { get; set; }
        public string? XId { get; set; }
        public string? MerchantLineItemNumber { get; set; }
        public string? BlankProviderNum { get; set; }
        public string? RbSku { get; set; }
        public string? Product { get; set; }
        public string? Size { get; set; }
        public string? BodyColor { get; set; }
        public int UnitShipped { get; set; }
        public decimal PrintCost { get; set; }
        public decimal HandlingCost { get; set; }
        public decimal BlankCost { get; set; }
        public decimal ShippingCost { get; set; }
        public decimal SurchargeCost { get; set; }
        public decimal Discount { get; set; }
        public decimal MiscCost { get; set; }
        public string? Currency { get; set; }
        public decimal TotalBilled { get; set; }
    }

    public class InvoiceRedBubbleForPackagingModel
    {
        public string? InvoiceDate { get; set; }
        public string? InvoiceNumber { get; set; }
        public string? FulfillerName { get; set; }
        public string? Facility { get; set; }
        public string? Packaging { get; set; }
        public string? PackagingSku { get; set; }
        public decimal PerItemCost { get; set; }
        public int Units { get; set; }
        public decimal TotalCost { get; set; }
    }
}
