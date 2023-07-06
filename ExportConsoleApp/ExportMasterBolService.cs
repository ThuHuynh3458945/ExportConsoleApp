using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace ExportConsoleApp
{
    public class ExportMasterBolService
    {
        #region Export Master Bol
        public async Task<byte[]> ExportMasterBolAsync(ExportMasterBolModel data)
        {
            if (data == null || data.Summary == null)
            {
                return null;
            }

            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                BindDataMasterBolSheet(xlPackage, data.FactoryId, data.Summary);

                foreach (var pallet in data.Pallets)
                {
                    BindDataMasterBolSheet(xlPackage, data.FactoryId, pallet);
                }

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }

        private void BindDataMasterBolSheet(ExcelPackage xlPackage, int factoryId, ExportMasterBolDataModel data)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add(data.SheetName);
            var row = 1;

            #region set width columns
            worksheet.Column(1).Width = 17.57d;
            worksheet.Column(2).Width = 24.43d;
            worksheet.Column(3).Width = 24.43d;
            worksheet.Column(5).Width = 38.29d;
            worksheet.Column(6).Width = 31.29d;
            worksheet.Column(7).Width = 19d;
            worksheet.Column(8).Width = 19d;
            #endregion set width columns

            #region Ship information (Left)
            worksheet.Cells[row, 1].Value = "SHIP FROM:";
            worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            worksheet.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1, row, 5].Merge = true;
            row++;

            worksheet.Cells[row, 1].Value = "TEE SHIRT CENTRAL, INC.";
            worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[row, 1, row, 5].Merge = true;
            row++;

            if (factoryId == (int)EFactory.Miami)
            {
                worksheet.Cells[row, 1].Value = "16085 NW 52nd Ave.";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "Miami Gardens, FL 33014";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;
            }
            else
            {
                worksheet.Cells[row, 1].Value = "PLANTA 31";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "AVE DE LOS TORRES # 2446";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "PARQUE INDUSTRIAL LOS BRAVOS";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "CIUDAD JUAREZ, CHIHUAHUA 32575, MEXICO";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;
            }

            worksheet.Cells[row, 1, row, 5].Merge = true;
            row++;

            worksheet.Cells[row, 1].Value = "SHIP TO:";
            worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            worksheet.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1, row, 5].Merge = true;
            row++;

            if (factoryId == (int)EFactory.Miami)
            {
                worksheet.Cells[row, 1].Value = "Pickup";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;
            }

            else
            {
                worksheet.Cells[row, 1].Value = "TECMA PAN AMERICAN WAREHOUSE";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "9571 PAN AMERICAN DRIVE";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1].Value = "EL PASO, TX 79927";
                worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;

                worksheet.Cells[row, 1, row, 5].Merge = true;
                row++;
            }
            worksheet.Cells[14, 1, 14, 5].Merge = true;
            row++;
            #endregion Ship information (Left)

            #region Total value information (Right)
            int rowValue = 1;
            worksheet.Cells[rowValue, 8].Value = "MOVEMENT #";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.MovementNumber;
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "BILL OF LADING #";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.BillOfLading.ToString();
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "TOTAL PACKAGES";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.TotalPackages;
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "TOTAL PALLETS/HAMPERS/BINS";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.TotalPallets;
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "TOTAL NET WEIGHT (LBS)";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.TotalNetWeight;
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "TOTAL GROSS WEIGHT (LBS)";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.TotalGrossWeight;
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rowValue++;

            worksheet.Cells[rowValue, 8].Value = "TOTAL VALUE";
            worksheet.Cells[rowValue, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[rowValue, 8, rowValue, 9].Merge = true;

            worksheet.Cells[rowValue, 10].Value = data.TotalValue;
            worksheet.Cells[rowValue, 10].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
            worksheet.Cells[rowValue, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var rangeShipTo = worksheet.Cells[1, 1, 6, 5];
            rangeShipTo.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rangeShipTo.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rangeShipTo.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rangeShipTo.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            var rangeShipFrom = worksheet.Cells[8, 1, 13, 5];
            rangeShipFrom.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rangeShipFrom.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rangeShipFrom.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rangeShipFrom.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            var rangeRight = worksheet.Cells[1, 10, 7, 10];
            rangeRight.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rangeRight.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rangeRight.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rangeRight.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            rangeShipTo.Style.Font.Name = "Arial";
            rangeShipTo.Style.Font.Size = 10;

            rangeShipFrom.Style.Font.Name = "Arial";
            rangeShipFrom.Style.Font.Size = 10;

            rangeRight.Style.Font.Name = "Arial";
            rangeRight.Style.Font.Size = 10;
            #endregion Total value information (Right)

            #region Title
            var properties = new List<string>
            {
                "PART #/GTIN #",
                "U.O.M.",
                "QTY PCS",
                "UNIT COST",
                "TOTAL VALUE",
                "NET WT. (LBS)",
                "GR. WT. (LBS)",
                "C.O.O.",
                "MEXICAN HTS CODE",
                "US HTS"
            };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = properties[i];

                worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                worksheet.Cells[row, i + 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells[row, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            row++;
            #endregion Title

            #region Bind Data
            foreach (var detail in data.Details)
            {
                int col = 1;
                worksheet.Cells[row, col].Value = detail.PartGtin;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.UOM;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.QtyPcs;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;


                worksheet.Cells[row, col].Value = detail.UnitCost;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, col].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                col++;

                worksheet.Cells[row, col].Value = detail.TotalValue;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[row, col].Style.Numberformat.Format = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* \"-\"??_-;_-@_-";
                col++;

                worksheet.Cells[row, col].Value = detail.TotalNetWeight;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.TotalGrossWeight;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.CooName;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.MexicanHtsCode;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                col++;

                worksheet.Cells[row, col].Value = detail.UsHts;
                worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row++;
            }
            #endregion Bind Data

            worksheet.Cells.AutoFitColumns();
        }
        #endregion Export Master Bol

        #region Export Section 321
        public async Task<byte[]> ExportSection321Async(List<ExportSection321Model> items)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                var worksheet = xlPackage.Workbook.Worksheets.Add("Tracking Status");
                var row = 1;

                #region Title
                var properties = new List<string>
                {
                    "Shipment Control Number",
                    "Shipment Type",
                    "Shipper Name",
                    "Shipper Address",
                    "Shipper City",
                    "Shipper Country",
                    "Shipper State",
                    "Shipper Postal",
                    "Shipper Port of Lading",
                    "Consignee Name",
                    "Consignee Address",
                    "Consignee City",
                    "Consignee Country",
                    "Consignee State",
                    "Consignee Postal",
                    "Product Description",
                    "Product Qty",
                    "Product UOM",
                    "Product Weight",
                    "Product Unit of Weight",
                    "Product Value",
                    "Cust. Reference",
                    "US Port Arr.",
                    "Fn Port Loading",
                    "Fn Port Reciept",
                    "Origin"
                };

                var yellowProperties = new[]
               {
                    "Shipment Control Number",
                    "Product Description",
                    "Origin"
                };

                for (var i = 0; i < properties.Count; i++)
                {
                    worksheet.Cells[row, i + 1].Value = properties[i];
                    worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    if (yellowProperties.Any(x => x == properties[i]))
                    {
                        worksheet.Cells[1, i + 1, 1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    }
                    else
                    {
                        worksheet.Cells[1, i + 1, 1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                    }
                    worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[row, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    worksheet.Cells[row, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                }
                row++;
                #endregion Title

                #region Bind Data
                foreach (var item in items)
                {
                    var col = 1;

                    worksheet.Cells[row, col].Value = item.ShipmentControlNumber;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipmentType;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperName;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperAdddress;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperCity;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperCountry;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperState;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperPostal;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ShipperPortofLading;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ConsigneeName;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ConsigneeAddress;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ConsigneeCity;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ConsigneeCountry;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    //TDP-1729: get full state/province name from table stateprovice.name for shipping_country=(us | ca)
                    var consigneeState = item.ConsigneeState;
                    //if (!string.IsNullOrEmpty(item.ConsigneeState)
                    //    && (item.ConsigneeCountry.ToLower() == "us" || item.ConsigneeCountry.ToLower() == "ca"))
                    //{
                    //    consigneeState = stateProvince.FirstOrDefault(x => x.Abbreviation.ToLower() == item.ConsigneeState.ToLower())?.Name;
                    //}
                    worksheet.Cells[row, col].Value = consigneeState;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ConsigneePostal;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductDescription;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductQty;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0";
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductUOM;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductWeight;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductUnitofWeight;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.ProductValue;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.CustReference;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.USPortArr;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.FnPortLoading;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.FnPortReceipt;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    col++;

                    worksheet.Cells[row, col].Value = item.Origin;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    
                    row++;
                }
                #endregion Bind Data

                worksheet.Cells[1, 1, row - 1, 26].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, row - 1, 26].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, row - 1, 26].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, row - 1, 26].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                worksheet.Cells.AutoFitColumns();
                worksheet.View.FreezePanes(2, 1);
                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
        #endregion Export Section 321
    }
}
