namespace ExportConsoleApp.Models
{
    public class InvoiceForReportModel
    {
        public string PartnerId { get; set; } = null!;
        public DateTime FromDate { get; set; }
        public DateTime EndDate { get; set; }

        public MdInvoiceForReportModel? MdInvoice { get; set; }
        public MdInvoiceForReportModel? TscInvoice { get; set; }
        public List<TscInvoiceDetailForReportModel> OtsInvoices { get; set; } = new List<TscInvoiceDetailForReportModel>();
        public List<AuthorizedCreditForReportModel> Credits { get; set; } = new List<AuthorizedCreditForReportModel>();
    }

    public class MdInvoiceForReportModel
    {
        public InvoiceTotalForReportModel Total { get; set; } = new InvoiceTotalForReportModel();
        public List<InvoiceDetailForReportModel> Details { get; set; } = new List<InvoiceDetailForReportModel>();
    }

    public class InvoiceTotalForReportModel
    {
        public int TotalOrders { get; set; }
        public int TotalQuantity { get; set; }
        public decimal TotalLine { get; set; }
        public decimal BagFee { get; set; }
        public decimal PackagingFee { get; set; }
        public decimal ShippingFee { get; set; }
        public decimal ShipCost { get; set; }
        public decimal TotalCreditAmount { get; set; }
    }

    public class InvoiceDetailForReportModel
    {
        public string InvoiceWeek { get; set; } = null!;
        public string PartnerId { get; set; } = null!;
        public string FactoryName { get; set; } = null!;
        public DateTime OrderDate { get; set; }
        public DateTime? ShipDate { get; set; }
        public DateTime? CancelledOnUtc { get; set; }
        public DateTime? CreditedOnUtc { get; set; }
        public int OrderId { get; set; }
        public string? PartnerOrderId { get; set; }
        public string? MiscOrderId { get; set; }
        public string TypeOrder { get; set; } = null!;
        public string? Upc { get; set; }
        public string? Sku { get; set; }
        public string? PartnerSku { get; set; }
        public string? PartnerBlankSku { get; set; }
        public string? License { get; set; }
        public string? PartnerItem { get; set; }
        public string? Description { get; set; }
        public string? SizeClassDesc { get; set; }
        public string? SizeDesc { get; set; }
        public string? ColorDesc { get; set; }
        public int Front { get; set; }
        public int Back { get; set; }
        public int Left { get; set; }
        public int Right { get; set; }
        public int Quantity { get; set; }
        public decimal FulfillmentUnitCost { get; set; }
        public decimal GarmentUnitCost { get; set; }
        public decimal TotalLine { get; set; }
        public decimal PackagingCost { get; set; }
        public decimal ShipCost { get; set; }
        public string? Type { get; set; }
        public string? Customer { get; set; }
        public string? Address1 { get; set; }
        public string? Address2 { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? ZipCode { get; set; }
        public string? Country { get; set; }
        public string? ShippingCarrier { get; set; }
        public string? ShippingPriority { get; set; }
        public string? TrackingNumber { get; set; }
        public string? Notes { get; set; }
    }

    public class AuthorizedCreditForReportModel
    {
        public string InvoiceWeek { get; set; } = null!;
        public string PartnerId { get; set; } = null!;
        public DateTime? CreditedOnUtc { get; set; }
        public int OrderId { get; set; }
        public string? PartnerOrderId { get; set; }
        public string? MiscOrderId { get; set; }
        public decimal CreditAmount { get; set; }
        public string? CreditReason { get; set; }
    }

    public class TscInvoiceDetailForReportModel
    {
        public string? MachineVendor { get; set; }
        public string? Company { get; set; }
        public string? PoNumber { get; set; }
        public DateTime CancelDate { get; set; }
        public string? Style { get; set; }
        public string? Description { get; set; }
        public string? Color { get; set; }
        public DateTime PrintDate { get; set; }
        public int PrintQty { get; set; }
        public decimal UnitCost { get; set; }
        public decimal InvoicePrice { get; set; }
    }
}
