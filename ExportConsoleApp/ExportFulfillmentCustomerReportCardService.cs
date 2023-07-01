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
        public FulfillmentCustomerReportCard BuildData()
        {
            var result = new FulfillmentCustomerReportCard();

            return result;
        }

        private static Table CreateHeader(string subject)
        {
            #region Define Table
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
                    .SetFontSize(7))
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
            paragraph.Add(new Text("")
                    .SetFont(fontNormal)
                    .SetFontSize(7))
                    .SetFontColor(ColorConstants.WHITE)
                    .SetBold()
                    ;

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
            var fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            var table = new Table(1);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Data
            //Blank Header
            var paragraph = new Paragraph();
            paragraph.Add(new Text("")
                    .SetFont(fontNormal)
                    .SetFontSize(4))
                    .SetFontColor(ColorConstants.BLACK)
                    .SetBold()
                ;

            var cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBorder(Border.NO_BORDER)
                ;
            table.AddCell(cell);

            //Header
            paragraph = new Paragraph();
            paragraph.Add(new Text(name)
                    .SetFont(fontNormal)
                    .SetFontSize(4))
                    .SetFontColor(ColorConstants.BLACK)
                    .SetTextAlignment(TextAlignment.CENTER)
                    .SetBold()
                ;

            cell = new Cell(1, 1);
            cell.Add(paragraph)
                .SetBackgroundColor(new DeviceRgb(217, 225, 242))
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .SetBorder(Border.NO_BORDER)
                ;
            table.AddCell(cell);

            //Blank Header
            paragraph = new Paragraph();
            paragraph.Add(new Text("")
                    .SetFont(fontNormal)
                    .SetFontSize(4))
                .SetFontColor(ColorConstants.WHITE)
                .SetBold()
                ;

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
            var table = new Table(18);
            table.SetWidth(UnitValue.CreatePercentValue(100f));
            table.SetFixedLayout()
                .SetBorder(Border.NO_BORDER);
            #endregion Define Table

            #region Bind Dat
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
            table.AddCell(AddCellTable(cell2Name, 4, 5));

            //Cell 3
            table.AddCell(AddCellTable(cell3Name, 4, 5));

            //Cell 4
            table.AddCell(AddCellTable(cell4Name, 4, 5));

            //cell empty
            table.AddCell(AddCellEmptyTable(18));
            #endregion Bind Data

            return table;
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
                ;

            return cell;
        }

        private static Cell AddCellEmptyTable(int colSpan)
        {
            var paragraph = new Paragraph();
            paragraph.Add(new Text(""));

            var cell = new Cell(1, colSpan);
            cell.Add(paragraph)
                .SetBorder(Border.NO_BORDER);

            return cell;
        }

        public byte[] ExportFulfillmentCustomerReportCard(string subject, FulfillmentCustomerReportCard model)
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
            var subHeaderRevenue = CreateSubHeader("Revenue");
            document.Add(subHeaderRevenue);

            var titleRevenue = CreateTitleTable("Revenue");
            document.Add(titleRevenue);
            #endregion Revenue 

            #endregion Bind Data

            document.Close();
            return stream.ToArray();
        }
    }
}
