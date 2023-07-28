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

    public class ExportFactoryReconciliationModel
    {
        public EFactory Factory { get; set; }

        public List<ExportReconciliationDetailModel> ReconciliationNonBlanks { get; set; }
        public List<ExportReceivingDetailModel> ReceivingNonBlanks { get; set; }

        public List<ExportReconciliationDetailModel> ReconciliationBlanks { get; set; }
        public List<ExportReceivingDetailModel> ReceivingBlanks { get; set; }

        public List<ExportAdjustDetailModel> RejectsAdj { get; set; }
        public List<ExportAdjustDetailModel> SamplesAdj { get; set; }
        public List<ExportAdjustDetailModel> DamagesAdj { get; set; }
        public List<ExportAdjustDetailModel> SubstitutionsAdj { get; set; }
        public List<ExportAdjustDetailModel> OtherAdj { get; set; }

        public List<ExportNegativeOnHandDetailModel> NegativeOnHands { get; set; }

        public ExportFactoryReconciliationModel()
        {
            ReconciliationNonBlanks = new List<ExportReconciliationDetailModel>();
            ReceivingNonBlanks = new List<ExportReceivingDetailModel>();

            ReconciliationBlanks = new List<ExportReconciliationDetailModel>();
            ReceivingBlanks = new List<ExportReceivingDetailModel>();

            RejectsAdj = new List<ExportAdjustDetailModel>();
            SamplesAdj = new List<ExportAdjustDetailModel>();
            DamagesAdj = new List<ExportAdjustDetailModel>();
            SubstitutionsAdj = new List<ExportAdjustDetailModel>();
            OtherAdj = new List<ExportAdjustDetailModel>();

            NegativeOnHands = new List<ExportNegativeOnHandDetailModel>();
        }
    }

    public class ExportReceivingDetailModel
    {
        public string? Vendor { get; set; }
        public DateTime TransactionDateOnUtc { get; set; }
        public string? RefNumber { get; set; }
        public int Qty { get; set; }
        public string? ReceiveAgainstPo { get; set; }
        public string? Items { get; set; }
        public string? ItemsDescription { get; set; }
        public decimal UnitCost { get; set; }
        public string? Memo { get; set; }
    }

    public class ExportAdjustDetailModel
    {
        public string? RefNumber { get; set; }
        public string? Item { get; set; }
        public int AdjQty { get; set; }
        public string? Account { get; set; }
        public string? TimePeriod { get; set; }
        public DateTime TransactionDateOnUtc { get; set; }
        public string? Memo { get; set; }
    }

    public class ExportNegativeOnHandDetailModel
    {
        public string? Sku { get; set; }
        public string? Inventory { get; set; }
        public int OnHand { get; set; }
    }
}
