namespace ExportConsoleApp.Models
{
    public class RoyaltyReportModel
    {
        public RoyaltyReportSummaryDataModel Summary { get; set; } = new RoyaltyReportSummaryDataModel();
        public List<RoyaltyReportSummaryDetailDataModel> SummaryDetails { get; set; } = new List<RoyaltyReportSummaryDetailDataModel>();
        public List<RoyaltyReportDetailDataModel> Details { get; set; } = new List<RoyaltyReportDetailDataModel>();
    }

    public class RoyaltyReportSummaryDataModel
    {
        public List<string> PartnerNames { get; set; } = new List<string>();
        public List<string> LicensorNames { get; set; } = new List<string>();
        public List<string> LicenseNames { get; set; } = new List<string>();
        public List<string> SubLicenseNames { get; set; } = new List<string>();
        public List<string> ArtistNames { get; set; } = new List<string>();
        public List<string> Groups { get; set; } = new List<string>();
        public List<string> Classes { get; set; } = new List<string>();
        public string? DesignId { get; set; }
        public DateTime? ShipDateFrom { get; set; }
        public DateTime? ShipDateTo { get; set; }

        public int TotalQty { get; set; }
        public decimal TotalRevenue { get; set; }
        public decimal TotalRoyalty { get; set; }
    }

    public class RoyaltyReportSummaryDetailDataModel
    {
        public string? PartnerArtId { get; set; }
        public string? PartnerName { get; set; }
        public string? LicensorName { get; set; }
        public decimal LicensorRoyalty { get; set; }
        public string? LicenseName { get; set; }
        public decimal LicenseRoyalty { get; set; }
        public string? SubLicenseName { get; set; }
        public decimal SubLicenseRoyalty { get; set; }
        public string? ArtistName { get; set; }
        public decimal ArtistRoyalty { get; set; }
        public string? Group { get; set; }
        public string? Class { get; set; }
        public string? VendorStyle { get; set; }
        public string? SizeName { get; set; }
        public string? ColorDesc { get; set; }
        public decimal RetailPrice { get; set; }
        public decimal Wholesale { get; set; }
        public string? StyleDescr { get; set; }
        public int Quantity { get; set; }
    }

    public class RoyaltyReportDetailDataModel
    {
        public string? PartnerArtId { get; set; }
        public string? PartnerName { get; set; }
        public string? LicensorName { get; set; }
        public string? LicensorRoyaltyType { get; set; }
        public decimal LicensorRoyalty { get; set; }
        public string? LicenseName { get; set; }
        public string? LicenseRoyaltyType { get; set; }
        public decimal LicenseRoyalty { get; set; }
        public string? SubLicenseName { get; set; }
        public string? SubLicenseRoyaltyType { get; set; }
        public decimal SubLicenseRoyalty { get; set; }
        public string? ArtistName { get; set; }
        public string? ArtistRoyaltyType { get; set; }
        public decimal ArtistRoyalty { get; set; }
        public string? Group { get; set; }
        public string? Class { get; set; }
        public string? VendorStyle { get; set; }
        public string? SizeName { get; set; }
        public string? ColorDesc { get; set; }
        public decimal RetailPrice { get; set; }
        public decimal Wholesale { get; set; }
        public int OrderId { get; set; }
        public string? Xid { get; set; }
        public string? StyleDescr { get; set; }
        public string? PartnerSku { get; set; }
        public string? Sku { get; set; }
        public int Quantity { get; set; }
        public DateTime? SortedOnUtc { get; set; }
        public string? ShippingCountry { get; set; }
    }
}
