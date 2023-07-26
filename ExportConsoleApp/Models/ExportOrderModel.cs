namespace ExportConsoleApp.Models
{
    public class ExportOrderSummaryModel
    {
        public int Id { get; set; }
        public int? Units { get; set; }
        public string? Partner { get; set; }
        public string? PartnerOrderId { get; set; }
        public DateTime? ReceivedOn { get; set; }
        public DateTime? ShippedOn { get; set; }
        public DateTime? SortedOn { get; set; }
        public DateTime? CancelledOnUtc { get; set; }
        public DateTime? CreditedOnUtc { get; set; }
        public DateTime? SlaDateOnUtc { get; set; }
        public DateTime? DroppedOnUtc { get; set; }
        public string? MergeBinId { get; set; }
        public DateTime? MergedOnUtc { get; set; }
        public string? Customer { get; set; }
        public string? FulfillmentStatus { get; set; }
        public string? CreditStatus { get; set; }
        public string? TrackingNumber { get; set; }
        public DateTime? PrintedOnUtc { get; set; }
        public string? OrderType { get; set; }
        public string? ShippingState { get; set; }
        public string? ShippingCountry { get; set; }
        public string? ShippingCarrier { get; set; }
        public string? FactoryName { get; set; }
        public string? EtaDateOnEst { get; set; }
        public bool IsPriority { get; set; }
    }

    public class ExportOrderDetailSummaryModel
    {
        public string? Type { get; set; }
        public int OrderId { get; set; }
        public string? Xid { get; set; }
        public DateTime ReceivedOnUtc { get; set; }
        public DateTime? CancelledOnUtc { get; set; }
        public DateTime? ShippedOnUtc { get; set; }
        public string? ItemName { get; set; }
        public string? Sku { get; set; }
        public string? GarmentId { get; set; }
        public string? VendorStyle { get; set; }
        public string? Color { get; set; }
        public string? Size { get; set; }
        public int Quantity { get; set; }
        public int HitCount { get; set; }
        public decimal? InvoiceUnitPrice { get; set; }
        public decimal? InvoiceTotal { get; set; }
        public string? FactoryName { get; set; }
    }
}
