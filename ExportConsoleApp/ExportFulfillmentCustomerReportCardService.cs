using ExportConsoleApp.Models;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace ExportConsoleApp
{
    public class ExportFulfillmentCustomerReportCardService
    {
        private static PdfFont FontNormal => PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
        private static PdfFont FontBold => PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

        public FulfillmentCustomerReportCardModel BuildData()
        {
            var result = new FulfillmentCustomerReportCardModel
            {
                Revenue = new FulfillmentCustomerReportCardDetailModel
                {
                    TopPartners = new List<FulfillmentCustomerReportCardDataModel>
                    {
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Printify",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "TeePublic",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "RedBubble",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "SenPrints",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Dreamship",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Canva",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Snow Commerce",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Gooten",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        }
                    },
                    OtherCustomer = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "All Other Customers",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Total = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Total",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Average = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Average Daily Revenue ",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = ""
                    }
                },
                Unit = new FulfillmentCustomerReportCardDetailModel
                {
                    TopPartners = new List<FulfillmentCustomerReportCardDataModel>
                    {
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Printify",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "TeePublic",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "RedBubble",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "SenPrints",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Dreamship",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Canva",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Snow Commerce",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Gooten",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        }
                    },
                    OtherCustomer = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "All Other Customers",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Total = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Total",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Average = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Average Daily Revenue ",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = ""
                    }
                },
                AverageSalePrice = new FulfillmentCustomerReportCardDetailModel
                {
                    TopPartners = new List<FulfillmentCustomerReportCardDataModel>
                    {
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Printify",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "TeePublic",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "RedBubble",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "SenPrints",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Dreamship",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Canva",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Snow Commerce",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        },
                        new FulfillmentCustomerReportCardDataModel
                        {
                            Name = "Gooten",
                            DailyActual = "$33",
                            DailyProjected = "$33",
                            DailyActualProjected = "$33",
                            DailyPriorYear = "$33",
                            DailyActualPriorYear = "$33",
                            WeeklyActual = "$33",
                            WeeklyProjected = "$33",
                            WeeklyActualProjected = "$33",
                            WeeklyPriorYear = "$33",
                            WeeklyActualPriorYear = "$33",
                            MonthlyActual = "$33",
                            MonthlyProjected = "$33",
                            MonthlyActualProjected = "$33",
                            MonthlyPriorYear = "$33",
                            MonthlyActualPriorYear = "$33"
                        }
                    },
                    OtherCustomer = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "All Other Customers",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Total = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Total",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "$33",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "$33",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "$33",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "$33",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "$33",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = "$33"
                    },
                    Average = new FulfillmentCustomerReportCardDataModel
                    {
                        Name = "Average Daily Revenue ",
                        DailyActual = "$33",
                        DailyProjected = "$33",
                        DailyActualProjected = "",
                        DailyPriorYear = "$33",
                        DailyActualPriorYear = "",
                        WeeklyActual = "$33",
                        WeeklyProjected = "$33",
                        WeeklyActualProjected = "",
                        WeeklyPriorYear = "$33",
                        WeeklyActualPriorYear = "",
                        MonthlyActual = "$33",
                        MonthlyProjected = "$33",
                        MonthlyActualProjected = "",
                        MonthlyPriorYear = "$33",
                        MonthlyActualPriorYear = ""
                    }
                }
            };

            return result;
        }

        private static Cell AddCellEmptyTable(int colSpan, bool isBorder = false)
        {
            var paragraph = new Paragraph();
            paragraph.Add(new Text(""));

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetBorder(isBorder ? new SolidBorder(0) : Border.NO_BORDER);

            return cell;
        }

        private static Cell AddCellTitleTable(string cellName, int colSpan, int frontSize, int width)
        {
            var table = new Table(colSpan);

            //Add cell name
            var text = new Text(cellName);
            text.SetFont(FontBold)
                .SetFontSize(frontSize)
                .SetFontColor(ColorConstants.BLACK)
                ;

            var paragraph = new Paragraph().Add(text);
            var cell = new Cell(1,3)
                .Add(paragraph)
                .SetWidth(width)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBorder(Border.NO_BORDER)
                .SetBorderBottom(new SolidBorder(0))
                ;
            table.AddCell(cell);

            var resultCell = new Cell(1, colSpan)
                    .Add(table)
                    .SetBorder(Border.NO_BORDER)
                ;

            return resultCell;
        }
        
        private static Cell AddCellValueTable(
            string cellName,
            int colSpan,
            int frontSize,
            TextAlignment textAlignment = TextAlignment.CENTER,
            bool isBold = false,
            bool isItalic = false,
            bool isBorder = false,
            float paddingLeft = 0,
            iText.Kernel.Colors.Color? backgroundColor = null)
        {
            var text = new Text(cellName);
            text.SetFont(isBold ? FontBold : FontNormal)
                .SetFontSize(frontSize)
                .SetFontColor(ColorConstants.BLACK);

            var paragraph = new Paragraph().Add(text);
            if (isItalic)
            {
                paragraph.SetItalic();
            }

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetTextAlignment(textAlignment)
                .SetBorder(Border.NO_BORDER)
                .SetBorderBottom(isBorder ? new SolidBorder(0) : Border.NO_BORDER)
                .SetPaddingLeft(paddingLeft)
                .SetPaddingBottom(5)
                ;

            if (backgroundColor != null)
            {
                cell.SetBackgroundColor(backgroundColor);
            }

            return cell;
        }

        private static void BindHeader(Document document, string subject)
        {
            #region Define Table
            const int font7 = 7;
            
            var table = new Table(1);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Header
            var paragraph = new Paragraph();
            paragraph.Add(new Text(subject)
                    .SetFont(FontBold)
                    .SetFontSize(font7))
                    .SetFontColor(ColorConstants.WHITE)
                    .SetTextAlignment(TextAlignment.CENTER)
                    ;

            var cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBackgroundColor(new DeviceRgb(68, 84, 106))
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetBorder(Border.NO_BORDER)
                ;
            table.AddCell(cell);

            //Blank Header
            paragraph = new Paragraph();
            paragraph.Add(new Text(""));

            cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBackgroundColor(new DeviceRgb(68, 84, 106))
                .SetBorderTop(new SolidBorder(ColorConstants.WHITE, 0))
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                ;
            table.AddCell(cell);
            #endregion Bind Data

            document.Add(table);
        }

        private static Table BindHeaderTable(string name)
        {
            #region Define Table
            const int font4 = 4;

            var fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var table = new Table(1);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Blank Header
            var cellEmpty = AddCellEmptyTable(1);
            table.AddCell(cellEmpty);

            //Header
            var paragraph = new Paragraph();
            paragraph.Add(new Text(name)
                    .SetFont(fontNormal)
                    .SetFontSize(font4))
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    
                ;

            var cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBackgroundColor(new DeviceRgb(217, 225, 242))
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetBorder(Border.NO_BORDER)
                ;
            table.AddCell(cell);

            //Blank Header
            paragraph = new Paragraph();
            paragraph.Add(new Text(""));

            cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBackgroundColor(new DeviceRgb(217, 225, 242))
                .SetBorderTop(new SolidBorder(0))
                .SetBorderBottom(Border.NO_BORDER)
                .SetBorderRight(Border.NO_BORDER)
                .SetBorderLeft(Border.NO_BORDER)
                ;
            table.AddCell(cell);
            #endregion Bind Data

            return table;
        }

        private static Table BindTitleTable(string title)
        {
            #region Define Table
            const int font4 = 4;

            var table = new Table(18);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            string cell2Name;
            string cell3Name;
            string cell4Name;

            if (title == "ASP")
            {
                cell2Name = "Daily ASP";
                cell3Name = "Weekly ASP";
                cell4Name = "Monthly ASP";
            }
            else
            {
                cell2Name = $"{title} / Day";
                cell3Name = $"{title} / Week";
                cell4Name = $"{title} / Month";
            }

            //Cell 1
            table.AddCell(AddCellEmptyTable(3));

            //Cell 2
            table.AddCell(AddCellTitleTable(cell2Name, 5, font4, 150));

            //Cell 3
            table.AddCell(AddCellTitleTable(cell3Name, 5, font4, 150));

            //Cell 4
            table.AddCell(AddCellTitleTable(cell4Name, 5, font4, 150));

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
        }

        private static Table BindSubTitleTable()
        {
            #region Define Table
            const int font4 = 4;

            var table = new Table(18);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Cell 1
            table.AddCell(AddCellEmptyTable(3));

            //Per title has 5 sub-tiles(Actual, Projected, vs. Proj, Prior Yr, vs. PY)
            for (int i = 1; i <= 3; i++)
            {
                //Actual
                table.AddCell(AddCellTitleTable("Actual", 1, font4,30));

                //Projected
                table.AddCell(AddCellTitleTable("Projected", 1, font4, 30));

                //vs. Proj
                table.AddCell(AddCellTitleTable("vs. Proj", 1, font4, 30));

                //Prior Yr
                table.AddCell(AddCellTitleTable("Prior Yr", 1, font4, 30));

                //vs. PY
                table.AddCell(AddCellTitleTable("vs. PY", 1, font4, 30));
            }

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
        }

        private static Table BindDataTable(
            FulfillmentCustomerReportCardDataModel data,
            TextAlignment textAlignment,
            bool isBorder = false,
            bool isBold = false,
            bool isItalic = false,
            float paddingLeft = 0,
            iText.Kernel.Colors.Color? backgroundColor = null)
        {
            #region Define Table
            const int font4 = 4;

            var table = new Table(18);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Name
            table.AddCell(AddCellValueTable(data.Name, 3,font4, textAlignment, isBold, isItalic, isBorder, paddingLeft, backgroundColor));

            #region Daily
            //Actual
            table.AddCell(AddCellValueTable(data.DailyActual, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Projected
            table.AddCell(AddCellValueTable(data.DailyProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. Proj
            table.AddCell(AddCellValueTable(data.DailyActualProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Prior Yr
            table.AddCell(AddCellValueTable(data.DailyPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. PY
            table.AddCell(AddCellValueTable(data.DailyActualPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));
            #endregion Daily

            #region Weekly
            //Actual
            table.AddCell(AddCellValueTable(data.WeeklyActual, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Projected
            table.AddCell(AddCellValueTable(data.WeeklyProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. Proj
            table.AddCell(AddCellValueTable(data.WeeklyActualProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Prior Yr
            table.AddCell(AddCellValueTable(data.WeeklyPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. PY
            table.AddCell(AddCellValueTable(data.WeeklyActualPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));
            #endregion Weekly

            #region Monthly
            //Actual
            table.AddCell(AddCellValueTable(data.MonthlyActual, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Projected
            table.AddCell(AddCellValueTable(data.MonthlyProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. Proj
            table.AddCell(AddCellValueTable(data.MonthlyActualProjected, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //Prior Yr
            table.AddCell(AddCellValueTable(data.MonthlyPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));

            //vs. PY
            table.AddCell(AddCellValueTable(data.MonthlyActualPriorYear, 1, font4, isBorder: isBorder, backgroundColor: backgroundColor));
            #endregion Monthly

            //cell empty
            if (data.Name != "All Other Customers" && data.Name != "Total")
            {
                table.AddCell(AddCellEmptyTable(18));
            }
            #endregion Bind Data

            return table;
        }

        private static void BindTable(
            Document document, 
            FulfillmentCustomerReportCardDetailModel data, 
            string name)
        {
            #region Title
            //Header
            var header = BindHeaderTable(name);
            document.Add(header);

            //Title
            var title = BindTitleTable(name);
            document.Add(title);

            //Sub-Title
            var subTitle = BindSubTitleTable();
            document.Add(subTitle);
            #endregion Title

            //Top Partner
            foreach (var partner in data.TopPartners)
            {
                var resp = BindDataTable(partner, TextAlignment.LEFT, isBorder: true, paddingLeft: 10);
                document.Add(resp);
            }

            //All Other Customers
            var otherCustomer= BindDataTable(data.OtherCustomer, TextAlignment.RIGHT, isBorder: true, isItalic: true);
            document.Add(otherCustomer);

            //Total
            var total = BindDataTable(data.Total, TextAlignment.LEFT, isBold: true, paddingLeft: 5, backgroundColor: new DeviceRgb(242, 242, 242));
            document.Add(total);

            //Average Daily
            if (name != "ASP")
            {
                var averageDaily = BindDataTable(data.Average, TextAlignment.LEFT, paddingLeft: 10, backgroundColor: new DeviceRgb(242, 242, 242));
                document.Add(averageDaily);
            }
        }

        public byte[] ExportFulfillmentCustomerReportCard(string subject, FulfillmentCustomerReportCardModel model)
        {
            using var stream = new MemoryStream();

            #region Define Document
            var writer = new PdfWriter(stream);
            var pdfDocument = new PdfDocument(writer);
            pdfDocument.SetDefaultPageSize(PageSize.LETTER);
            var document = new Document(pdfDocument);
            document.SetMargins(35f, 35f, 35f, 35f);
            #endregion Define Document

            #region Bind Data

            //Header 
            BindHeader(document, subject);

            //Revenue
            BindTable(document, model.Revenue, "Revenue");

            //Units
            BindTable(document, model.Unit, "Units");

            //Average Sale Price
            BindTable(document, model.AverageSalePrice, "ASP");

            #endregion Bind Data

            document.Close();
            return stream.ToArray();
        }
    }
}
