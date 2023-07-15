namespace ExportConsoleApp.Models
{
    public class InvoicePrintifyModel
    {
        public string? Xid { get; set; }
        public string? OrderId { get; set; }
        public string? Status { get; set; }
        public string? DtgSkuOrdered { get; set; }
        public string? DtgSku { get; set; }
        public DateTime ReceivedOnUtc { get; set; }
        public DateTime? SortedOnUtc { get; set; }
        public int ShipQty { get; set; }
        public decimal BlankCosts { get; set; }
        public decimal PrintCosts { get; set; }
        public int FrontPrint { get; set; }
        public int BackPrint { get; set; }
        public string? SleevePrints { get; set; }
        public string? InnerLabel { get; set; }
        public string? OuterLabel { get; set; }
        public decimal HandlingCosts { get; set; }
        public decimal Discount { get; set; }
        public decimal TotalProductCosts { get; set; }
        public int TaxRate { get; set; }
        public decimal Tax { get; set; }
        public decimal TotalProductCostWithTax { get; set; }
        public string? ShipState { get; set; }
        public string? ShippingCountry { get; set; }
        public string? TrackingNumber { get; set; }
    }
}
