namespace ExportConsoleApp.Models
{
    public class FulfillmentCustomerReportCardModel
    {
        public FulfillmentCustomerReportCardDetailModel Revenue { get; set; } = new FulfillmentCustomerReportCardDetailModel();
        public FulfillmentCustomerReportCardDetailModel Unit { get; set; } = new FulfillmentCustomerReportCardDetailModel();
        public FulfillmentCustomerReportCardDetailModel AverageSalePrice { get; set; } = new FulfillmentCustomerReportCardDetailModel();
    }

    public class FulfillmentCustomerReportCardDetailModel
    {
        public List<FulfillmentCustomerReportCardDataModel> TopPartners { get; set; } = new List<FulfillmentCustomerReportCardDataModel>();
        public FulfillmentCustomerReportCardDataModel OtherCustomer { get; set; } = new FulfillmentCustomerReportCardDataModel();
        public FulfillmentCustomerReportCardDataModel Total { get; set; } = new FulfillmentCustomerReportCardDataModel();
        public FulfillmentCustomerReportCardDataModel Average { get; set; } = new FulfillmentCustomerReportCardDataModel();
    }

    public class FulfillmentCustomerReportCardDataModel
    {
        public string Name { get; set; } = null!;

        #region Daily
        public string DailyActual { get; set; }= null!;
        public string DailyProjected { get; set; } = null!;
        //vs. Proj = (Actual - Projected) / Actual
        public string DailyActualProjected { get; set; } = null!;
        public string DailyPriorYear { get; set; } = null!;
        //vs. PY = (Actual - Prior Yr) / Actual
        public string DailyActualPriorYear { get; set; } = null!;
        #endregion Daily

        #region Weekly
        public string WeeklyActual { get; set; } = null!;
        public string WeeklyProjected { get; set; } = null!;
        //vs. Proj = (Actual - Projected) / Actual
        public string WeeklyActualProjected { get; set; } = null!;
        public string WeeklyPriorYear { get; set; } = null!;
        //vs. PY = (Actual - Prior Yr) / Actual
        public string WeeklyActualPriorYear { get; set; } = null!;
        #endregion Weekly

        #region Monthly
        public string MonthlyActual { get; set; } = null!;
        public string MonthlyProjected { get; set; } = null!;
        //vs. Proj = (Actual - Projected) / Actual
        public string MonthlyActualProjected { get; set; } = null!;
        public string MonthlyPriorYear { get; set; } = null!;
        //vs. PY = (Actual - Prior Yr) / Actual
        public string MonthlyActualPriorYear { get; set; } = null!;
        #endregion Monthly
    }
}
