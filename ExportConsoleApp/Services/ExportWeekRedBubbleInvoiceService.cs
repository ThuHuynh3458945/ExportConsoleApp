using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportConsoleApp.Services
{
    public class ExportWeekRedBubbleInvoice
    {
        public InvoiceRedBubbleModel BuildData()
        {
            return new InvoiceRedBubbleModel
            {
                Garments = new List<InvoiceRedBubbleForGarmentModel>
                {
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "63408176",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    },
                    new InvoiceRedBubbleForGarmentModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "sandiego-ca-us",
                        XId = "",
                        MerchantLineItemNumber = "269744223",
                        BlankProviderNum = "BC_3001",
                        RbSku = "D9OE6",
                        Product = "mens-t-shirt",
                        Size = "large",
                        BodyColor = "red",
                        UnitShipped = 1,
                        PrintCost = (decimal)3.75,
                        HandlingCost = (decimal)0.1,
                        BlankCost = (decimal)3.09,
                        ShippingCost = (decimal)0,
                        SurchargeCost = (decimal)0.09,
                        Discount = (decimal)0,
                        MiscCost = (decimal)0,
                        Currency = "USD",
                        TotalBilled = (decimal)7.03,
                    }
                },
                Packaging = new List<InvoiceRedBubbleForPackagingModel>
                {
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    },
                    new InvoiceRedBubbleForPackagingModel
                    {
                        InvoiceDate = "2023-06-18",
                        InvoiceNumber = "",
                        FulfillerName = "monsterdigital",
                        Facility = "elpaso-tx-us",
                        Packaging = "PolyMailer Medium",
                        PackagingSku = "RBPKG007",
                        PerItemCost = (decimal)0.3,
                        Units = 1742,
                        TotalCost = (decimal)522.6
                    }
                }
            };
        }

        private static void BindGarmentsSheet(ExcelPackage xlPackage, List<InvoiceRedBubbleForGarmentModel> items)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Garments");
            var row = 1;

            #region Title
            var properties = new List<string>
                {
                    "invoice_date",
                    "invoice_number",
                    "fulfiller_name",
                    "facility",
                    "order_id",
                    "item_id",
                    "blank_provider_num",
                    "rb_sku",
                    "product",
                    "size",
                    "body_color",
                    "units_shipped",
                    "1. print_cost",
                    "2. handling_cost",
                    "3. blank_cost",
                    "4. shipping_cost",
                    "5. surcharge_cost",
                    "6. discount",
                    "7. misc_cost",
                    "currency",
                    "total_billed"
                };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1, 1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            row++;
            #endregion Title

            #region Bind Data Detail

            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.InvoiceDate;
                col++;

                worksheet.Cells[row, col].Value = item.InvoiceNumber;
                col++;

                worksheet.Cells[row, col].Value = item.FulfillerName;
                col++;

                worksheet.Cells[row, col].Value = item.Facility;
                col++;

                worksheet.Cells[row, col].Value = item.XId;
                col++;

                worksheet.Cells[row, col].Value = item.MerchantLineItemNumber;
                col++;

                worksheet.Cells[row, col].Value = item.BlankProviderNum;
                col++;

                worksheet.Cells[row, col].Value = item.RbSku;
                col++;

                worksheet.Cells[row, col].Value = item.Product;
                col++;

                worksheet.Cells[row, col].Value = item.Size;
                col++;

                worksheet.Cells[row, col].Value = item.BodyColor;
                col++;

                worksheet.Cells[row, col].Value = item.UnitShipped;
                col++;

                worksheet.Cells[row, col].Value = item.PrintCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.HandlingCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.BlankCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.ShippingCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.SurchargeCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Discount;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.MiscCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Currency;
                col++;

                worksheet.Cells[row, col].Value = item.TotalBilled;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";

                row++;
            }
            #endregion Bind Data Detail

            var range = worksheet.Cells[row - items.Count, 1, row - 1, properties.Count];
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            worksheet.Cells.AutoFitColumns(0);
        }

        private static void BindPackagingSheet(ExcelPackage xlPackage, List<InvoiceRedBubbleForPackagingModel> items)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Packaging");
            var row = 1;

            #region Title
            var properties = new List<string>
                {
                    "invoice_date",
                    "invoice_number",
                    "fulfiller_name",
                    "facility",
                    "packaging",
                    "RB_packaging_sku",
                    "per_item_cost",
                    "units",
                    "total_cost"
                };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1, 1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[1, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            row++;
            #endregion Title

            #region Bind Data Detail

            foreach (var item in items)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.InvoiceDate;
                col++;

                worksheet.Cells[row, col].Value = item.InvoiceNumber;
                col++;

                worksheet.Cells[row, col].Value = item.FulfillerName;
                col++;

                worksheet.Cells[row, col].Value = item.Facility;
                col++;

                worksheet.Cells[row, col].Value = item.Packaging;
                col++;

                worksheet.Cells[row, col].Value = item.PackagingSku;
                col++;

                worksheet.Cells[row, col].Value = item.PerItemCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";
                col++;

                worksheet.Cells[row, col].Value = item.Units;
                col++;

                worksheet.Cells[row, col].Value = item.TotalCost;
                worksheet.Cells[row, col].Style.Numberformat.Format = "$#,##0.00";

                row++;
            }
            #endregion Bind Data Detail

            #region Bind Data Total
            worksheet.Cells[row, properties.Count - 1].Value = "Total";
            worksheet.Cells[row, properties.Count - 1].Style.Font.Bold = true;
            worksheet.Cells[row, properties.Count - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Cells[row, properties.Count].Value = items.Sum(o => o.TotalCost);
            worksheet.Cells[row, properties.Count].Style.Font.Bold = true;
            worksheet.Cells[row, properties.Count].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[row, properties.Count].Style.Numberformat.Format = "$#,##0.00";

            #endregion Bind Data Total

            var range = worksheet.Cells[row - items.Count, 1, row - 1, properties.Count];
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            worksheet.Cells.AutoFitColumns(0);
        }

        public async Task<byte[]> ExportWeekRedBubbleInvoiceAsync(InvoiceRedBubbleModel data)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                if (data.Garments.Any())
                {
                    BindGarmentsSheet(xlPackage, data.Garments);
                }

                if (data.Packaging.Any())
                {
                    BindPackagingSheet(xlPackage, data.Packaging);
                }

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
