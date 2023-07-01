namespace ExportConsoleApp.Models
{
    public class PrintPalletScanFormModel
    {
        public int PalletId { get; set; }
        public string? ClosedBy { get; set; }
        public string ClosedOn { get; set; }
        public int SkidNumber { get; set; }
        public string? SKidName { get; set; }
        public List<PrintPalletScanFormDetailModel> Details { get; set; } = new List<PrintPalletScanFormDetailModel>();
    }

    public class PrintPalletScanFormDetailModel
    {
        public string PartnerName { get; set; } = null!;
        public string TrackingCodes { get; set; } = null!;
        public string ScanFormUrl { get; set; } = null!;
        public string CarrierCode { get; set; } = null!;
    }
}
