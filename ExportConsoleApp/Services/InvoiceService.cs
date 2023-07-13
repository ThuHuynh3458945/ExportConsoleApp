using ExportConsoleApp.Heplers;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using ExportConsoleApp.Models;
using System.Drawing;

namespace ExportConsoleApp.Services
{
    public class InvoiceService
    {
        public InvoiceForReportModel BuildData()
        {
            return new InvoiceForReportModel
            {

            };
        }

        private static void BindInvoicesSheet(ExcelPackage xlPackage, 
            List<InvoiceDetailForReportModel> invoicesDetails,
            InvoiceTotalForReportModel total)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Invoices");
            var row = 1;

            #region Title
            var properties = new List<string>
                {
                    "Invoice Week",
                    "Partner Id",
                    "Factory",
                    "Order Date",
                    "ShipDate",
                    "Cancel Date",
                    "Credited Date",
                    "Order ID",
                    "Partner Order ID",
                    "Misc Order ID",
                    "UPC",
                    "SKU",
                    "Partner SKU",
                    "Partner Blank SKU",
                    "License",
                    "Partner Item",
                    "Style Description",
                    "Size Class",
                    "Size",
                    "Color",
                    "Front",
                    "Back",
                    "Left",
                    "Right",
                    "Quantity",
                    "Fulfillment Unit Cost",
                    "Garment Unit Cost",
                    "Line Total",
                    "Packaging Cost",
                    "Ship Cost",
                    "Credit Amount",
                    "Fulfillment Type",
                    "Customer",
                    "Address Line 1",
                    "Address Line 2",
                    "City",
                    "State",
                    "ZipCode",
                    "Country",
                    "Shipping Carrier",
                    "Shipping Priority",
                    "TrackingNumber",
                    "Notes",
                    "Credit Reason"
                };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            row++;
            #endregion Title

            #region Bind Data Detail

            foreach (var invoice in invoicesDetails)
            {
                var col = 1;
                //InvoiceWeek
                worksheet.Cells[row, col].Value = invoice.InvoiceWeek;
                col++;

                //PartnerId
                worksheet.Cells[row, col].Value = invoice.PartnerId;
                col++;

                //Factory
                worksheet.Cells[row, col].Value = invoice.FactoryName;
                col++;

                //OrderDate
                worksheet.Cells[row, col].Value = DateTimeHelper.ConvertToEstTime(invoice.OrderDate).ToString("M/d/yyyy h:mm tt");
                col++;

                //ShipDate
                worksheet.Cells[row, col].Value = invoice.ShipDate.HasValue ? DateTimeHelper.ConvertToEstTime(invoice.ShipDate.Value).ToString("M/d/yyyy h:mm tt") : "";
                col++;

                //Cancel Date
                worksheet.Cells[row, col].Value = invoice.CancelledOnUtc.HasValue ? DateTimeHelper.ConvertToEstTime(invoice.CancelledOnUtc.Value).ToString("M/d/yyyy h:mm tt") : "";
                col++;

                //Credited Date
                worksheet.Cells[row, col].Value = invoice.CreditedOnUtc.HasValue ? DateTimeHelper.ConvertToEstTime(invoice.CreditedOnUtc.Value).ToString("M/d/yyyy h:mm tt") : "";
                col++;

                //Order ID
                worksheet.Cells[row, col].Value = invoice.OrderId;
                col++;

                //Partner Order ID
                worksheet.Cells[row, col].Value = invoice.PartnerOrderId;
                col++;

                //Misc Order ID
                worksheet.Cells[row, col].Value = invoice.MiscOrderId;
                col++;

                //UPC
                worksheet.Cells[row, col].Value = invoice.TypeOrder == "SAMPLE" ? "SAMPLE" : invoice.Upc;
                col++;

                //SKU
                worksheet.Cells[row, col].Value = invoice.Sku;
                col++;

                //PartnerSku
                worksheet.Cells[row, col].Value = invoice.PartnerSku;
                col++;

                //Partner Blank SKU
                worksheet.Cells[row, col].Value = invoice.PartnerBlankSku;
                col++;

                //License
                worksheet.Cells[row, col].Value = invoice.License;
                col++;

                //Partner Item
                worksheet.Cells[row, col].Value = invoice.PartnerItem;
                col++;

                //Style Description
                worksheet.Cells[row, col].Value = invoice.Description;
                col++;

                //Size Class
                worksheet.Cells[row, col].Value = invoice.SizeClassDesc;
                col++;

                //Size
                worksheet.Cells[row, col].Value = invoice.SizeDesc;
                col++;

                //Color
                worksheet.Cells[row, col].Value = invoice.ColorDesc;
                col++;

                //Front
                worksheet.Cells[row, col].Value = invoice.Front > 0 ? invoice.Front : string.Empty;
                col++;

                //Back
                worksheet.Cells[row, col].Value = invoice.Back > 0 ? invoice.Back : string.Empty;
                col++;

                //Left
                worksheet.Cells[row, col].Value = invoice.Left;
                col++;

                //Right
                worksheet.Cells[row, col].Value = invoice.Right;
                col++;

                //Quantity
                worksheet.Cells[row, col].Value = invoice.Quantity;
                col++;

                //Fulfillment Unit Cost
                //TODO-Minh: check TFS to set correct FulfillmentUnitCost
                worksheet.Cells[row, col].Value = invoice.FulfillmentUnitCost;
                col++;

                //Garment Unit Cost
                //TODO-Minh: check TFS to set correct GarmentUnitCost
                worksheet.Cells[row, col].Value = invoice.GarmentUnitCost;
                col++;

                //Line Total
                //TODO-Minh: check TFS to set correct TotalLine
                worksheet.Cells[row, col].Value = invoice.TotalLine;
                col++;

                //Packaging Cost
                worksheet.Cells[row, col].Value = invoice.PackagingCost;
                col++;

                //Ship Cost
                worksheet.Cells[row, col].Value = Math.Round(invoice.ShipCost, 2);
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                //Credit Amount
                worksheet.Cells[row, col].Value = 0;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                //Fulfillment Type
                worksheet.Cells[row, col].Value = invoice.Type == "ScreenPrint" ? "Screen Print" : "Digital Print";
                col++;

                //Customer
                worksheet.Cells[row, col].Value = (invoice.PartnerId == "hmcreator" || invoice.PartnerId == "hmcs") ? string.Empty : invoice.Customer;
                col++;

                //Address Line 1
                worksheet.Cells[row, col].Value = (invoice.PartnerId == "hmcreator" || invoice.PartnerId == "hmcs") ? string.Empty : invoice.Address1;
                col++;

                //Address Line 2
                worksheet.Cells[row, col].Value = (invoice.PartnerId == "hmcreator" || invoice.PartnerId == "hmcs") ? string.Empty : invoice.Address2;
                col++;

                //City
                worksheet.Cells[row, col].Value = invoice.City;
                col++;

                //State
                worksheet.Cells[row, col].Value = invoice.State;
                col++;

                //ZipCode
                worksheet.Cells[row, col].Value = invoice.ZipCode;
                col++;

                //Country
                worksheet.Cells[row, col].Value = invoice.Country;
                col++;

                //Shipping Carrier
                worksheet.Cells[row, col].Value = invoice.ShippingCarrier;
                col++;

                //Shipping Priority
                worksheet.Cells[row, col].Value = invoice.ShippingPriority;
                col++;
                //TrackingNumber
                worksheet.Cells[row, col].Value = invoice.TrackingNumber;
                col++;

                //Notes
                worksheet.Cells[row, col].Value = invoice.Notes;

                row++;
            }
            #endregion Bind Data Detail

            #region Bind Data Sub Total
            worksheet.Cells[row, 8].Value = total.TotalOrders;
            worksheet.Cells[row, 25].Value = total.TotalQuantity;
            worksheet.Cells[row, 28].Value = total.TotalLine;
            worksheet.Cells[row, 29].Value = total.PackagingFee;

            worksheet.Cells[row, 30].Value = total.ShipCost;
            worksheet.Cells[row, 30].Style.Numberformat.Format = "0.00";

            worksheet.Cells[row, 31].Value = total.TotalCreditAmount;
            worksheet.Cells[row, 31].Style.Numberformat.Format = "0.00";

            row += 2;//space between sub total and total.
            #endregion Bind Data Sub Total


            #region Bind Data Total
            //Total Orders
            worksheet.Cells[row, 1].Value = "Total Orders";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.TotalOrders;
            row++;

            //Total Units
            worksheet.Cells[row, 1].Value = "Total Units";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.TotalQuantity;
            row++;

            //Total line
            worksheet.Cells[row, 1].Value = "Line Total";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.TotalLine;
            row++;

            //Bag Fee
            worksheet.Cells[row, 1].Value = "Bag Fee";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.BagFee;
            row++;

            worksheet.Cells[row, 1].Value = "Packaging Fee";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.PackagingFee;
            row++;

            worksheet.Cells[row, 1].Value = "Shipping Fee";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.ShippingFee;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            worksheet.Cells[row, 1].Value = "Credit Total";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.TotalCreditAmount;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            //Invoice Total                    
            worksheet.Cells[row, 1].Value = "Invoice Total";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = total.TotalLine + total.ShippingFee + total.PackagingFee + total.BagFee + total.TotalCreditAmount;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";

            worksheet.Cells.AutoFitColumns(0);
            #endregion Bind Data Total

            worksheet.View.FreezePanes(2, 1);
        }

        public async Task<byte[]> ExportInvoiceAsync(InvoiceForReportModel model)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                if (!model.Invoices.Any() || model.InvoiceTotal == null)
                {
                    var worksheet = xlPackage.Workbook.Worksheets.Add("Invoices");

                    const int row = 2;
                    const int col = 1;

                    worksheet.Cells[row, col].Value = $"No orders were shipped from {model.FromDate:MM/dd/yyyy} through {model.EndDate:MM/dd/yyyy}";
                    worksheet.Cells[row, col].Style.Font.Size = 16;
                    worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
                    worksheet.Cells[row, col].Style.Font.Bold = true;
                }
                else
                {
                    BindInvoicesSheet(xlPackage, model.Invoices, model.InvoiceTotal);
                }
                
                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
