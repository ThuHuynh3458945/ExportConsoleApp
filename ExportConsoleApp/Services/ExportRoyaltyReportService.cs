using ExportConsoleApp.Helpers;
using ExportConsoleApp.Heplers;
using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportConsoleApp.Services
{
    public class ExportRoyaltyReportService
    {
        public RoyaltyReportModel BindData()
        {
            var result = new RoyaltyReportModel
            {
                Summary = new RoyaltyReportSummaryDataModel
                {
                    PartnerNames = new List<string>(),
                    LicensorNames = new List<string>(),
                    LicenseNames = new List<string>(),
                    SubLicenseNames = new List<string>(),
                    ArtistNames = new List<string>(),
                    Groups = new List<string>(),
                    Classes = new List<string>(),
                    DesignId = string.Empty,
                    ShipDateFrom = DateTime.Parse("2023-07-01 10:39:38.547"),
                    ShipDateTo = DateTime.Parse("2023-07-31 10:39:38.547"),
                    TotalQty = 1000,
                    TotalRevenue = 1000,
                    TotalRoyalty = 1000
                },
                SummaryDetails = new List<RoyaltyReportSummaryDetailDataModel>
                {
                    new RoyaltyReportSummaryDetailDataModel
                    {
                        PartnerArtId = "10431011",
                        PartnerName = "Kohls",
                        LicensorName = "Bravado",
                        LicensorRoyalty = (decimal)5.4375,
                        LicenseName = "BEASTIE BOYS",
                        LicenseRoyalty = (decimal)5.4375,
                        SubLicenseName = "",
                        SubLicenseRoyalty = (decimal)5.4375,
                        ArtistName = "",
                        ArtistRoyalty = (decimal)5.4375,
                        Group = "MENS",
                        Class = "TEE",
                        VendorStyle = "2168",
                        SizeName = "XL",
                        ColorDesc = "Natural Mineral Wash",
                        RetailPrice = (decimal)29.99,
                        Wholesale = (decimal)14.5,
                        StyleDescr = "BEASTIE BOYS COLOR LOGO Mens Natural Mineral Wash T shirt",
                        Quantity = 1
                    }, 
                    new RoyaltyReportSummaryDetailDataModel
                    {
                        PartnerArtId = "10431011",
                        PartnerName = "Kohls",
                        LicensorName = "Bravado",
                        LicensorRoyalty = (decimal)5.4375,
                        LicenseName = "BEASTIE BOYS",
                        LicenseRoyalty = (decimal)5.4375,
                        SubLicenseName = "",
                        SubLicenseRoyalty = (decimal)5.4375,
                        ArtistName = "",
                        ArtistRoyalty = (decimal)5.4375,
                        Group = "MENS",
                        Class = "TEE",
                        VendorStyle = "2168",
                        SizeName = "XL",
                        ColorDesc = "Natural Mineral Wash",
                        RetailPrice = (decimal)29.99,
                        Wholesale = (decimal)14.5,
                        StyleDescr = "BEASTIE BOYS COLOR LOGO Mens Natural Mineral Wash T shirt",
                        Quantity = 1
                    }
                },
                Details = new List<RoyaltyReportDetailDataModel>
                {
                    new RoyaltyReportDetailDataModel
                    {
                        PartnerArtId = "10431011",
                        PartnerName = "Kohls",
                        LicensorName = "Bravado",
                        LicensorRoyaltyType = "37.50% of Invoice",
                        LicensorRoyalty = (decimal)5.4375,
                        LicenseName = "BEASTIE BOYS",
                        LicenseRoyaltyType = "",
                        LicenseRoyalty = (decimal)5.4375,
                        SubLicenseName = "",
                        SubLicenseRoyaltyType = "",
                        SubLicenseRoyalty = (decimal)5.4375,
                        ArtistName = "",
                        ArtistRoyaltyType = "",
                        ArtistRoyalty = (decimal)5.4375,
                        Group = "MENS",
                        Class = "TEE",
                        VendorStyle = "2168",
                        SizeName = "XL",
                        ColorDesc = "Natural Mineral Wash",
                        RetailPrice = (decimal)29.99,
                        Wholesale = (decimal)14.5,
                        StyleDescr = "BEASTIE BOYS COLOR LOGO Mens Natural Mineral Wash T shirt",
                        Quantity = 1
                    },
                    new RoyaltyReportDetailDataModel
                    {
                        PartnerArtId = "10431011",
                        PartnerName = "Kohls",
                        LicensorName = "Bravado",
                        LicensorRoyaltyType = "37.50% of Invoice",
                        LicensorRoyalty = (decimal)5.4375,
                        LicenseName = "BEASTIE BOYS",
                        LicenseRoyaltyType = "",
                        LicenseRoyalty = (decimal)5.4375,
                        SubLicenseName = "",
                        SubLicenseRoyaltyType = "",
                        SubLicenseRoyalty = (decimal)5.4375,
                        ArtistName = "",
                        ArtistRoyaltyType = "",
                        ArtistRoyalty = (decimal)5.4375,
                        Group = "MENS",
                        Class = "TEE",
                        VendorStyle = "2168",
                        SizeName = "XL",
                        ColorDesc = "Natural Mineral Wash",
                        RetailPrice = (decimal)29.99,
                        Wholesale = (decimal)14.5,
                        StyleDescr = "BEASTIE BOYS COLOR LOGO Mens Natural Mineral Wash T shirt",
                        Quantity = 1
                    }
                }
            };

            return result;
        }

        private static void BindSummarySheet(ExcelPackage xlPackage, RoyaltyReportSummaryDataModel data)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Summary");
            int row = 1;

            #region Add Parameters
            var paramHeaderRange = worksheet.Cells[row, 1, 1, 2];
            paramHeaderRange.Value = "Report Parameters";
            paramHeaderRange.Style.Font.Bold = true;
            paramHeaderRange.Merge = true;
            paramHeaderRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            paramHeaderRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            paramHeaderRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            row++;

            worksheet.Cells[row, 1].Value = "Retailer";
            worksheet.Cells[row, 2].Value = data.PartnerNames.Any() ? data.PartnerNames.JoinComma(true) : "All Retailers";
            row++;

            worksheet.Cells[row, 1].Value = "Licensor";
            worksheet.Cells[row, 2].Value = data.LicensorNames.Any() ? data.LicensorNames.JoinComma(true) : "All Licensors";
            row++;

            worksheet.Cells[row, 1].Value = "License";
            worksheet.Cells[row, 2].Value = data.LicenseNames.Any() ? data.LicenseNames.JoinComma(true) : "All Licenses";
            row++;

            worksheet.Cells[row, 1].Value = "Sub-License";
            worksheet.Cells[row, 2].Value = data.SubLicenseNames.Any() ? data.SubLicenseNames.JoinComma(true) : "All Sub-Licenses";
            row++;

            worksheet.Cells[row, 1].Value = "Artist";
            worksheet.Cells[row, 2].Value = data.ArtistNames.Any() ? data.ArtistNames.JoinComma(true) : "All Artists";
            row++;

            var shipDateValue = data.ShipDateFrom?.ToString("MM/dd/yyyy");
            if (data.ShipDateFrom.HasValue && !data.ShipDateTo.HasValue)
            {
                shipDateValue += " - " + DateTime.UtcNow.ToString("MM/dd/yyyy");
            }
            else if (!data.ShipDateFrom.HasValue && data.ShipDateTo.HasValue)
            {
                shipDateValue += $"<= {data.ShipDateTo.Value.ToString("MM/dd/yyyy")}";
            }
            else if (data.ShipDateFrom.HasValue && data.ShipDateTo.HasValue)
            {
                shipDateValue += $" - {data.ShipDateTo.Value.ToString("MM/dd/yyyy")}";
            }
            worksheet.Cells[row, 1].Value = "Ship Date";
            worksheet.Cells[row, 2].Value = shipDateValue;
            row++;

            worksheet.Cells[row, 1].Value = "Style Group";
            worksheet.Cells[row, 2].Value = data.Groups.Any() ? data.Groups.JoinComma(true) : "All Group Styles";
            row++;

            worksheet.Cells[row, 1].Value = "Style Category";
            worksheet.Cells[row, 2].Value = data.Classes.Any() ? data.Classes.JoinComma(true) : "All Style Categories";
            row++;

            worksheet.Cells[row, 1].Value = "Design ID";
            worksheet.Cells[row, 2].Value = data.DesignId;
            row++;
            row++;

            var paramTableRange = worksheet.Cells[1, 1, 10, 2];
            paramTableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            paramTableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            paramTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            paramTableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            #endregion Add Parameters

            #region Add Summary
            var summaryHeaderRange = worksheet.Cells[row, 1, row, 3];
            summaryHeaderRange.Value = "Summary";
            summaryHeaderRange.Style.Font.Bold = true;
            summaryHeaderRange.Merge = true;
            summaryHeaderRange.Style.Font.Size = 12;
            summaryHeaderRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
            summaryHeaderRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            summaryHeaderRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            row++;

            int col = 1;
            worksheet.Cells[row, col].Value = "Qty";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Value = data.TotalQty;
            worksheet.Cells[row + 1, col].Style.Numberformat.Format = "#,##0";
            col++;

            worksheet.Cells[row, col].Value = "Total Revenue";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Value = data.TotalRevenue;
            worksheet.Cells[row + 1, col].Style.Numberformat.Format = "$#,##0.00";
            col++;

            worksheet.Cells[row, col].Value = "Total Royalty";
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells[row + 1, col].Value = data.TotalRoyalty;
            worksheet.Cells[row + 1, col].Style.Numberformat.Format = "$#,##0.00";

            var summaryTableRange = worksheet.Cells[row - 1, 1, row + 1, 3];
            summaryTableRange.Style.Font.Name = "Arial";
            summaryTableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            summaryTableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            summaryTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            summaryTableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            summaryTableRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            #endregion Add Summary

            worksheet.Column(1).Width = 25;
            worksheet.Column(2).Width = 25;
            worksheet.Column(3).Width = 25;
        }

        private static void BindSummaryDetailSheet(ExcelPackage xlPackage, List<RoyaltyReportSummaryDetailDataModel> items)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Summary Detail");
            int row = 1;

            #region Bind Title
            var titles = new List<string>
            {
                "Partner Art #",
                "Retailer",
                "Licensor",
                "Licensor Royalty",
                "License",
                "License Royalty",
                "Sub-License",
                "Sub-License Royalty ",
                "Artist",
                "Artist Royalty",
                "Group",
                "Category",
                "Garment Style #",
                "Size",
                "Color",
                "Retail Price",
                "Wholesale Price",
                "Description",
                "Qty"
            };
            
            for (int i = 0; i < titles.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = titles[i];
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            }
            worksheet.View.FreezePanes(2, 1);
            row++;
            #endregion Bind Title

            #region Bind Data

            #region Bind Items
            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.PartnerArtId;
                col++;

                worksheet.Cells[row, col].Value = item.PartnerName;
                col++;

                worksheet.Cells[row, col].Value = item.LicensorName;
                col++;

                worksheet.Cells[row, col].Value = item.LicensorRoyalty;
                worksheet.Cells[row , col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.LicenseName;
                col++;

                worksheet.Cells[row, col].Value = item.LicenseRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.SubLicenseName;
                col++;

                worksheet.Cells[row, col].Value = item.SubLicenseRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.ArtistName;
                col++;

                worksheet.Cells[row, col].Value = item.ArtistRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Group;
                col++;

                worksheet.Cells[row, col].Value = item.Class;
                col++;

                worksheet.Cells[row, col].Value = item.VendorStyle;
                col++;

                worksheet.Cells[row, col].Value = item.SizeName;
                col++;

                worksheet.Cells[row, col].Value = item.ColorDesc;
                col++;

                worksheet.Cells[row, col].Value = item.RetailPrice;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;
                
                worksheet.Cells[row, col].Value = item.StyleDescr;
                col++;

                worksheet.Cells[row, col].Value = item.Wholesale;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Quantity;

                row++;
            }
            #endregion Bind Items

            #region Bind Total
            var colTotal = 1;

            worksheet.Cells[row, colTotal].Value = "Totals";
            colTotal += 3;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.LicensorRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 2;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.LicenseRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 2;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.SubLicenseRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 2;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.ArtistRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 6;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.RetailPrice);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Wholesale);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal +=2 ;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Quantity);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";

            var toRow = items.Any() ? items.Count + 2 : 2;
            var rangeTotal = worksheet.Cells[row, 1, toRow, titles.Count];
            rangeTotal.Style.Font.Bold = true;
            rangeTotal.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rangeTotal.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(170, 170, 170));
            #endregion Bind Total

            var range = worksheet.Cells[1, 1, toRow, titles.Count];
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            #endregion Bind Data

            worksheet.Cells.AutoFitColumns();
        }

        private static void BindDetailSheet(ExcelPackage xlPackage, List<RoyaltyReportDetailDataModel> items)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Detail");
            int row = 1;

            #region Bind Title
            var titles = new List<string>
            {
                "Partner Art ",
                "Retailer",
                "Licensor",
                "Licensor Royalty Type",
                "Licensor Royalty $",
                "License",
                "License Royalty Type",
                "License Royalty $",
                "Sub-License",
                "Sub-License Royalty Type",
                "Sub-License Royalty $",
                "Artist",
                "Artist Royalty Type",
                "Artist Royalty $",
                "Group",
                "Category",
                "Garment Style #",
                "Size",
                "Color",
                "Retail Price",
                "Wholesale Price",
                "MD Order #",
                "Partner Order #",
                "Item Description",
                "Partner SKU",
                "MD SKU",
                "Qty",
                "Shipped On",
                "Ship To Country"
            };

            for (int i = 0; i < titles.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = titles[i];
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            }
            worksheet.View.FreezePanes(2, 1);
            row++;
            #endregion Bind Title

            #region Bind Data

            #region Bind Items
            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.PartnerArtId;
                col++;

                worksheet.Cells[row, col].Value = item.PartnerName;
                col++;

                worksheet.Cells[row, col].Value = item.LicensorName;
                col++;

                worksheet.Cells[row, col].Value = item.LicensorRoyaltyType;
                col++;

                worksheet.Cells[row, col].Value = item.LicensorRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.LicenseName;
                col++;

                worksheet.Cells[row, col].Value = item.LicenseRoyaltyType;
                col++;

                worksheet.Cells[row, col].Value = item.LicenseRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.SubLicenseName;
                col++;

                worksheet.Cells[row, col].Value = item.SubLicenseRoyaltyType;
                col++;

                worksheet.Cells[row, col].Value = item.SubLicenseRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.ArtistName;
                col++;

                worksheet.Cells[row, col].Value = item.ArtistRoyaltyType;
                col++;

                worksheet.Cells[row, col].Value = item.ArtistRoyalty;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Group;
                col++;

                worksheet.Cells[row, col].Value = item.Class;
                col++;

                worksheet.Cells[row, col].Value = item.VendorStyle;
                col++;

                worksheet.Cells[row, col].Value = item.SizeName;
                col++;

                worksheet.Cells[row, col].Value = item.ColorDesc;
                col++;

                worksheet.Cells[row, col].Value = item.RetailPrice;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Wholesale;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.OrderId;
                col++;

                worksheet.Cells[row, col].Value = item.Xid;
                col++;

                worksheet.Cells[row, col].Value = item.StyleDescr;
                col++;

                worksheet.Cells[row, col].Value = item.PartnerSku;
                col++;

                worksheet.Cells[row, col].Value = item.Sku;
                col++;

                worksheet.Cells[row, col].Value = item.Quantity;

                worksheet.Cells[row, col].Value = !item.SortedOnUtc.HasValue
                    ? string.Empty
                    : DateTimeHelper.ConvertToEstTime(item.SortedOnUtc.Value).ToString("MM/dd/yyyy");
                col++;

                worksheet.Cells[row, col].Value = item.ShippingCountry;

                row++;
            }
            #endregion Bind Items

            #region Bind Total
            var colTotal = 1;

            worksheet.Cells[row, colTotal].Value = "Totals";
            colTotal += 4;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.LicensorRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 3;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.LicenseRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 3;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.SubLicenseRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 3;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.ArtistRoyalty);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 6;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.RetailPrice);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Wholesale);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";
            colTotal += 6;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Quantity);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$#,##0.00";

            var toRow = items.Any() ? items.Count + 2 : 2;
            var rangeTotal = worksheet.Cells[row, 1, toRow, titles.Count];
            rangeTotal.Style.Font.Bold = true;
            rangeTotal.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rangeTotal.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(170, 170, 170));
            #endregion Bind Total

            var range = worksheet.Cells[1, 1, toRow, titles.Count];
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            #endregion Bind Data

            worksheet.Cells.AutoFitColumns();
        }

        public async Task<byte[]> ExportRoyaltyReportAsync(RoyaltyReportModel data)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                //Bind Summary sheet
                BindSummarySheet(xlPackage, data.Summary);
                //Bind Summary Detail sheet
                BindSummaryDetailSheet(xlPackage, data.SummaryDetails);
                //Bind Detail sheet
                BindDetailSheet(xlPackage, data.Details);

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
