namespace ExportConsoleApp.Models
{
    public class ExportSummaryReconciliationModel
    {
        public List<ExportSummaryReconciliationSummaryModel> Summaries { get; set; }
        public List<ExportReconciliationDetailModel> MiamiDetails { get; set; }
        public List<ExportReconciliationDetailModel> JuarezDetails { get; set; }
        public List<ExportReconciliationDetailModel> TijuanaDetails { get; set; }
        public ExportSummaryReconciliationModel()
        {
            Summaries = new List<ExportSummaryReconciliationSummaryModel>();
            MiamiDetails = new List<ExportReconciliationDetailModel>();
            JuarezDetails = new List<ExportReconciliationDetailModel>();
            TijuanaDetails = new List<ExportReconciliationDetailModel>();
        }
    }

    public class ExportSummaryReconciliationSummaryModel
    {
        public string? Factory { get; set; }
        public string? ItemType { get; set; }
        public int Units { get; set; }
        public decimal Price { get; set; }
    }

    public class ExportReconciliationDetailModel
    {
        public string? Sku { get; set; }
        public int BeginningQty { get; set; }
        public int Picked { get; set; }
        public int Received { get; set; }
        public int ManualIncrements { get; set; }
        public int ManualDecrements { get; set; }
        public int BegPlusTrans { get; set; }
        public int EndingQty { get; set; }
        public int Variance { get; set; }
        public int Sold { get; set; }
        public int Rejects { get; set; }
        public int Wip { get; set; }
        public decimal UnitCost { get; set; }
        public decimal SoldTotalCost { get; set; }
        public decimal AvgCost { get; set; }
        //EndingValue = AvgCost > 0 ? EndingQty * AvgCost : EndingQty * UnitCost;
        public decimal EndingValue { get; set; }
        public bool IsLockerStock { get; set; }
    }
}
