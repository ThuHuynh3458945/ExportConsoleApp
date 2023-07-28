using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportConsoleApp.Services
{
    public class ExportWeeklyReconciliationService
    {
        public ExportSummaryReconciliationModel BindData()
        {
            var result = new ExportSummaryReconciliationModel
            {
                Summaries = new List<ExportSummaryReconciliationSummaryModel>
                {
                    new ExportSummaryReconciliationSummaryModel
                    {
                        Factory = "Miami",
                        ItemType = "Regular",
                        Units = 706745,
                        Price = (decimal)2477274.98
                    },
                    new ExportSummaryReconciliationSummaryModel
                    {
                        Factory = "Juarez",
                        ItemType = "Regular",
                        Units = 808665,
                        Price = (decimal)3175516.88
                    },
                    new ExportSummaryReconciliationSummaryModel
                    {
                        Factory = "Juarez",
                        ItemType = "Locker Stock",
                        Units = 278426,
                        Price = (decimal)587745.73
                    },
                    new ExportSummaryReconciliationSummaryModel
                    {
                        Factory = "Tijuana",
                        ItemType = "Regular",
                        Units = 499273,
                        Price = (decimal)1719261.58
                    },
                },
                MiamiDetails = new List<ExportReconciliationDetailModel>
                {
                    new ExportReconciliationDetailModel
                    {
                        Sku = "093-154-005",
                        BeginningQty = 0,
                        Picked = 0,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 0,
                        EndingQty = 0,
                        Variance = 0,
                        Sold = 0,
                        Rejects = 0,
                        Wip = 0,
                        UnitCost = (decimal)14.23,
                        SoldTotalCost = 0,
                        AvgCost = (decimal)14.23,
                        EndingValue = 0,
                        IsLockerStock = false
                    },
                    new ExportReconciliationDetailModel
                    {
                        Sku = "093-128-001",
                        BeginningQty = 0,
                        Picked = 0,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 0,
                        EndingQty = 0,
                        Variance = 0,
                        Sold = 0,
                        Rejects = 0,
                        Wip = 0,
                        UnitCost = (decimal)14.23,
                        SoldTotalCost = 0,
                        AvgCost = (decimal)14.23,
                        EndingValue = 0,
                        IsLockerStock = false
                    }
                },
                JuarezDetails = new List<ExportReconciliationDetailModel>
                {
                    new ExportReconciliationDetailModel
                    {
                        Sku = "016-041-004",
                        BeginningQty = 279,
                        Picked = 8,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 271,
                        EndingQty = 271,
                        Variance = 0,
                        Sold = 8,
                        Rejects = 0,
                        Wip = 1,
                        UnitCost = (decimal)4.67,
                        SoldTotalCost = (decimal)37.36,
                        AvgCost = (decimal)4.93,
                        EndingValue = (decimal)1336.03,
                        IsLockerStock = false
                    },
                    new ExportReconciliationDetailModel
                    {
                        Sku = "016-041-005",
                        BeginningQty = 345,
                        Picked = 3,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 342,
                        EndingQty = 342,
                        Variance = 0,
                        Sold = 2,
                        Rejects = 0,
                        Wip = 1,
                        UnitCost = (decimal)5.6,
                        SoldTotalCost = (decimal)11.2,
                        AvgCost = (decimal)5.6,
                        EndingValue = (decimal)1915.2,
                        IsLockerStock = true
                    }
                },
                TijuanaDetails = new List<ExportReconciliationDetailModel>
                {
                    new ExportReconciliationDetailModel
                    {
                        Sku = "093-154-005",
                        BeginningQty = 0,
                        Picked = 0,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 0,
                        EndingQty = 0,
                        Variance = 0,
                        Sold = 0,
                        Rejects = 0,
                        Wip = 0,
                        UnitCost = (decimal)14.23,
                        SoldTotalCost = 0,
                        AvgCost = (decimal)14.23,
                        EndingValue = 0,
                        IsLockerStock = false
                    },
                    new ExportReconciliationDetailModel
                    {
                        Sku = "093-128-001",
                        BeginningQty = 0,
                        Picked = 0,
                        Received = 0,
                        ManualIncrements = 0,
                        ManualDecrements = 0,
                        BegPlusTrans = 0,
                        EndingQty = 0,
                        Variance = 0,
                        Sold = 0,
                        Rejects = 0,
                        Wip = 0,
                        UnitCost = (decimal)14.23,
                        SoldTotalCost = 0,
                        AvgCost = (decimal)14.23,
                        EndingValue = 0,
                        IsLockerStock = false
                    }
                }
            };

            return result;
        }

        private static void BindSummarySheet(
            ExcelPackage xlPackage, 
            List<ExportSummaryReconciliationSummaryModel> items
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Summary");
            var row = 1;

            #region Bind Title
            var titles = new List<string>
            {
                "Factory", 
                "Item Type", 
                "Units", 
                "$"
            };
            
            for (int i = 0; i < titles.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = titles[i];
                worksheet.Cells[row, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
            }
            worksheet.View.FreezePanes(2, 1);
            row++;
            #endregion Bind Title

            #region Bind Data
            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.Factory;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = item.ItemType;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = item.Units;
                worksheet.Cells[row, col].Style.Numberformat.Format = "###,###,##0";
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = item.Price;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$###,###,##0.00";
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                row++;
            }
            #endregion Bind Data

            #region Bind Total
            var colTotal = 1;
            var rangeTotal = worksheet.Cells[row, colTotal, row, colTotal + 1];
            rangeTotal.Value = "Totals";
            rangeTotal.Merge = true;
            rangeTotal.Style.Font.Bold = true;
            rangeTotal.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            colTotal++;
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Units);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "###,###,##0";
            worksheet.Cells[row, colTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Price);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "$###,###,##0.00";
            worksheet.Cells[row, colTotal].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            #endregion Bind Total

            var range = worksheet.Cells[1, 1, items.Count + 2, titles.Count];
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            worksheet.Cells.AutoFitColumns();
        }

        private static void BindReconciliationSheet(
            ExcelPackage xlPackage,
            string sheetName,
            bool isIncludeIsLockerStock,
            List<ExportReconciliationDetailModel> items
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add(sheetName);
            var row = 1;

            #region Bind Title
            var titles = new List<string>
            {
                "sku",
                "BeginningQty",
                "Picked",
                "Received",
                "ManualIncrements",
                "ManualDecrements",
                "BegPlusTrans",
                "EndingQty",
                "Variance",
                "Sold",
                "Rejects",
                "Wip",
                "UnitCost",
                "SoldTotalCost",
                "AvgCost",
                "EndingValue"
            };

            if (isIncludeIsLockerStock)
            {
                titles.Add("Is Locker Stock");
            }

            for (int i = 0; i < titles.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = titles[i];
                worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
            }
            worksheet.View.FreezePanes(2, 1);
            row++;
            #endregion Bind Title

            #region Bind Data
            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.Sku;
                col++;

                worksheet.Cells[row, col].Value = item.BeginningQty;
                col++;

                worksheet.Cells[row, col].Value = item.Picked;
                col++;

                worksheet.Cells[row, col].Value = item.Received;
                col++;

                worksheet.Cells[row, col].Value = item.ManualIncrements;
                col++;

                worksheet.Cells[row, col].Value = item.ManualDecrements;
                col++;

                worksheet.Cells[row, col].Value = item.BegPlusTrans;
                col++;

                worksheet.Cells[row, col].Value = item.EndingQty;
                col++;

                worksheet.Cells[row, col].Value = item.Variance;
                col++;

                worksheet.Cells[row, col].Value = item.Sold;
                col++;

                worksheet.Cells[row, col].Value = item.Rejects;
                col++;

                worksheet.Cells[row, col].Value = item.Wip;
                col++;

                worksheet.Cells[row, col].Value = item.UnitCost > 0 ? item.UnitCost : item.AvgCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                worksheet.Cells[row, col].Value = item.SoldTotalCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                worksheet.Cells[row, col].Value = item.AvgCost > 0 ? item.AvgCost : item.UnitCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                worksheet.Cells[row, col].Value = item.EndingValue;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;

                if (isIncludeIsLockerStock)
                {
                    worksheet.Cells[row, col].Value = item.IsLockerStock ? "Y" : "N";
                }
                row++;
            }
            #endregion Bind Data

            #region Bind Total
            var colTotal = 1;

            worksheet.Cells[row, colTotal].Value = "Totals";
            worksheet.Cells[row, colTotal].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, colTotal].Style.Font.Bold = true;
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.BeginningQty);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Picked);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Received);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.ManualIncrements);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.ManualDecrements);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.BegPlusTrans);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.EndingQty);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Variance);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Sold);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Rejects);
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.Wip);
            colTotal++;

            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.SoldTotalCost);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "0.00";
            colTotal++;
            colTotal++;

            worksheet.Cells[row, colTotal].Value = items.Sum(x => x.EndingValue);
            worksheet.Cells[row, colTotal].Style.Numberformat.Format = "0.00";
            #endregion Bind Total

            worksheet.Cells.AutoFitColumns();
        }

        public async Task<byte[]> ExportWeeklySummaryReconciliationAsync(ExportSummaryReconciliationModel data)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                //Summary sheet
                BindSummarySheet(xlPackage, data.Summaries);

                //Miami sheet
                BindReconciliationSheet(xlPackage, "Miami", false, data.MiamiDetails);

                //Juarez sheet
                BindReconciliationSheet(xlPackage, "Juarez", true, data.JuarezDetails);

                //Tijuana sheet
                BindReconciliationSheet(xlPackage, "Tijuana", false, data.TijuanaDetails);

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
