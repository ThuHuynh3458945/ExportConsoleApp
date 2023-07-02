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
                    }
                }
            };

            return result;
        }

        private static Cell AddCellEmptyTable(int colSpan, bool isSetBorder = false)
        {
            var paragraph = new Paragraph();
            paragraph.Add(new Text(""));

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetBorder(isSetBorder ? new SolidBorder(0) : Border.NO_BORDER);

            return cell;
        }

        private static Cell AddCellDataNameTable(string cellName, TextAlignment textAlignment , int frontSize, int colSpan)
        {
            var fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            var paragraph = new Paragraph();
            paragraph.Add(new Text(cellName)
                    .SetFont(fontNormal)
                    .SetFontSize(frontSize))
                    .SetFontColor(ColorConstants.BLACK)
                    .SetBold()
                    ;

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(textAlignment)
                .SetBorderBottom(new SolidBorder(0))
                .SetPaddingLeft(10f)
                ;

            return cell;
        }

        private static Cell AddCellTable(string cellName, int frontSize, int colSpan)
        {
            var fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

            var paragraph = new Paragraph();
            paragraph.Add(new Text(cellName)
                    .SetFont(fontNormal)
                    .SetFontSize(frontSize))
                    .SetFontColor(ColorConstants.BLACK)
                    .SetBold()
                    ;

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetBorder(Border.NO_BORDER)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBorderBottom(new SolidBorder(0))
                .SetPaddingBottom(5)
                ;

            return cell;
        }

        private static Table CreateHeader(string subject)
        {
            #region Define Table
            const int font7 = 7;

            var fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var table = new Table(1);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Header
            var paragraph = new Paragraph();
            paragraph.Add(new Text(subject)
                    .SetFont(fontNormal)
                    .SetFontSize(font7))
                    .SetFontColor(ColorConstants.WHITE)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBold()
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

            return table;
        }
        
        private static Table CreateSubHeader(string name)
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
                    .SetBold()
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

        private static Table CreateTitleTable(string title)
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
            table.AddCell(AddCellTable(cell2Name, font4, 5));

            //Cell 3
            table.AddCell(AddCellTable(cell3Name, font4, 5));

            //Cell 4
            table.AddCell(AddCellTable(cell4Name, font4, 5));

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
        }

        private static Table CreateSubTitleTable()
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
                table.AddCell(AddCellTable("Actual", font4, 1));

                //Projected
                table.AddCell(AddCellTable("Projected", font4, 1));

                //vs. Proj
                table.AddCell(AddCellTable("vs. Proj", font4, 1));

                //Prior Yr
                table.AddCell(AddCellTable("Prior Yr", font4, 1));

                //vs. PY
                table.AddCell(AddCellTable("vs. PY", font4, 1));
            }

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
        }

        private static Table BindDataTable(FulfillmentCustomerReportCardDataModel data, TextAlignment textAlignment)
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
            table.AddCell(AddCellDataNameTable(data.Name, textAlignment, font4, 3));

            #region Daily
            //Actual
            table.AddCell(AddCellTable(data.DailyActual, font4, 1));

            //Projected
            table.AddCell(AddCellTable(data.DailyProjected, font4, 1));

            //vs. Proj
            table.AddCell(AddCellTable(data.DailyActualProjected, font4, 1));

            //Prior Yr
            table.AddCell(AddCellTable(data.DailyPriorYear, font4, 1));

            //vs. PY
            table.AddCell(AddCellTable(data.DailyActualPriorYear, font4, 1));
            #endregion Daily

            #region Weekly
            //Actual
            table.AddCell(AddCellTable(data.WeeklyActual, font4, 1));

            //Projected
            table.AddCell(AddCellTable(data.WeeklyProjected, font4, 1));

            //vs. Proj
            table.AddCell(AddCellTable(data.WeeklyActualProjected, font4, 1));

            //Prior Yr
            table.AddCell(AddCellTable(data.WeeklyPriorYear, font4, 1));

            //vs. PY
            table.AddCell(AddCellTable(data.WeeklyActualPriorYear, font4, 1));
            #endregion Weekly

            #region Monthly
            //Actual
            table.AddCell(AddCellTable(data.MonthlyActual, font4, 1));

            //Projected
            table.AddCell(AddCellTable(data.MonthlyProjected, font4, 1));

            //vs. Proj
            table.AddCell(AddCellTable(data.MonthlyActualProjected, font4, 1));

            //Prior Yr
            table.AddCell(AddCellTable(data.MonthlyPriorYear, font4, 1));

            //vs. PY
            table.AddCell(AddCellTable(data.MonthlyActualPriorYear, font4, 1));
            #endregion Monthly

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
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

            #region Header 
            var header = CreateHeader(subject);
            document.Add(header);
            #endregion Header

            #region Revenue

            #region Title
            var subHeaderRevenue = CreateSubHeader("Revenue");
            document.Add(subHeaderRevenue);

            var titleRevenue = CreateTitleTable("Revenue");
            document.Add(titleRevenue);

            var subTitleRevenue = CreateSubTitleTable();
            document.Add(subTitleRevenue);
            #endregion Title

            #region Bind Data
            foreach (var partner in model.Revenue.TopPartners)
            {
                var data = BindDataTable(partner, TextAlignment.LEFT);
                document.Add(data);
            }

            // All Other Customers
            var otherCustomerRevenue = BindDataTable(model.Revenue.OtherCustomer, TextAlignment.RIGHT);
            document.Add(otherCustomerRevenue);
            #endregion Bind Data

            #endregion Revenue 

            #endregion Bind Data

            document.Close();
            return stream.ToArray();
        }
    }
}
