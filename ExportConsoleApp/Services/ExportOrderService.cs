using ExportConsoleApp.Enums;
using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportConsoleApp.Services
{
    public class ExportOrderService
    {
        public List<ExportOrderSummaryModel> BindDataOrder()
        {
            var result = new List<ExportOrderSummaryModel>
            {
                new ExportOrderSummaryModel
                {
                    Id = 4752772,
                    Units = 100,
                    Partner = "test",
                    PartnerOrderId =  "test",
                    ReceivedOn =  DateTime.UtcNow,
                    ShippedOn =  DateTime.UtcNow,
                    SortedOn = DateTime.UtcNow,
                    CancelledOnUtc =  DateTime.UtcNow,
                    CreditedOnUtc =  DateTime.UtcNow,
                    SlaDateOnUtc =  DateTime.UtcNow,
                    DroppedOnUtc = DateTime.UtcNow,
                    MergeBinId =  "test",
                    MergedOnUtc =  DateTime.UtcNow,
                    Customer =  "test",
                    FulfillmentStatus =  "test",
                    CreditStatus =  "test",
                    TrackingNumber =  "test",
                    PrintedOnUtc =  DateTime.UtcNow,
                    OrderType =  "test",
                    ShippingState =  "test",
                    ShippingCountry =  "test",
                    ShippingCarrier =  "test",
                    FactoryName =  "test",
                    EtaDateOnEst =  "test",
                    IsPriority = true
                }
            };

            return result;
        }

        public List<ExportOrderDetailSummaryModel> BindDataOrderDetail()
        {
            var result = new List<ExportOrderDetailSummaryModel>
            {
                new ExportOrderDetailSummaryModel
                {
                    Type = "test",
                    OrderId = 4752772,
                    Xid = "test",
                    ReceivedOnUtc = DateTime.UtcNow,
                    CancelledOnUtc = DateTime.UtcNow,
                    ShippedOnUtc = DateTime.UtcNow,
                    ItemName = "test",
                    Sku = "test",
                    GarmentId = "test",
                    VendorStyle = "test",
                    Color = "test",
                    Size = "test",
                    Quantity = 1,
                    HitCount = 1,
                    InvoiceUnitPrice = 100,
                    InvoiceTotal = 100,
                    FactoryName = "Miami"
                }
            };

            return result;
        }
        
        private static void AutoFitColumnExportExcel(ExcelWorksheet worksheet, int row, int col, Dictionary<int, int> listCheckCol)
        {
            if (!string.IsNullOrEmpty(worksheet.Cells[row, col].GetValue<string>())
                && worksheet.Cells[row, col].GetValue<string>().Length > listCheckCol[col])
            {
                listCheckCol[col] = worksheet.Cells[row, col].GetValue<string>().Length;
                worksheet.Cells[row, col].AutoFitColumns();
            }
        }

        private static List<string> GetColumnsOrderByOrderStatus(EOrderStatusExtend orderStatus)
        {
            var result = new List<string>();
            switch (orderStatus)
            {
                case EOrderStatusExtend.All:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "DroppedOn",
                        "MergeBinId",
                        "MergedOn",
                        "ShippedOn",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date",
                        "TrackingNumber",
                    });
                    break;
                case EOrderStatusExtend.Error:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.AlertAdmin:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.ReprintRequests:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.OnHold:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.Unassigned:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "OrderType",
                        "Priority",
                        "Items",
                        "Factory",
                        "ReceivedOn",
                        "Customer",
                        "CreditedOn",
                        "CreditStatus"
                    });
                    break;
                case EOrderStatusExtend.Unshipped:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "DroppedOn",
                        "MergedOn",
                        "MergeBinId",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.InShortage:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.NotDropped:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.InProgress:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "SortedOn",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.Packaged:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "PackagedOn",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
                case EOrderStatusExtend.Shipped:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "SlaDate",
                        "SortedOn",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date",
                        "TrackingNumber"
                    });
                    break;
                case EOrderStatusExtend.Canceled:
                    result.AddRange(new List<string>()
                    {
                        "Id",
                        "PartnerId",
                        "PartnerOrderId",
                        "Factory",
                        "OrderType",
                        "Priority",
                        "Items",
                        "ReceivedOn",
                        "CancelledOn",
                        "Customer",
                        "FulfillmentStatus",
                        "ETA Date"
                    });
                    break;
            }
            return result;
        }

        private static void BindOrderSheet(
            ExcelPackage xlPackage,
            List<string> columns, 
            List<ExportOrderSummaryModel> orders
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Orders");

            #region Bind Title
            var col = 0;
            var row = 1;

            var properties = new string[columns.Count];

            if (columns.Contains("Id"))
            {
                properties[col] = "Order #";
                col++;
            }

            if (columns.Contains("PartnerId"))
            {
                properties[col] = "Partner ID";
                col++;
            }

            if (columns.Contains("PartnerOrderId"))
            {
                properties[col] = "Partner Order ID";
                col++;
            }
            if (columns.Contains("TscId"))
            {
                properties[col] = "TSC Order Id";
                col++;
            }
            if (columns.Contains("Factory"))
            {
                properties[col] = "Factory Name";
                col++;
            }
            if (columns.Contains("OrderType"))
            {
                properties[col] = "Order Type";
                col++;
            }
            if (columns.Contains("Priority"))
            {
                properties[col] = "Priority";
                col++;
            }
            if (columns.Contains("Items"))
            {
                properties[col] = "Item Count";
                col++;
            }

            if (columns.Contains("ReceivedOn"))
            {
                properties[col] = "Received On";
                col++;
            }

            if (columns.Contains("SlaDate"))
            {
                properties[col] = "Sla Date";
                col++;
            }

            if (columns.Contains("DroppedOn"))
            {
                properties[col] = "Dropped On";
                col++;
            }

            if (columns.Contains("MergeBinId"))
            {
                properties[col] = "Merge Bin Id";
                col++;
            }

            if (columns.Contains("MergedOn"))
            {
                properties[col] = "Merged On";
                col++;
            }

            if (columns.Contains("ShippedOn"))
            {
                properties[col] = "Shipped On";
                col++;
            }

            if (columns.Contains("SortedOn"))
            {
                properties[col] = "Shipped On";
                col++;
            }

            if (columns.Contains("CancelledOn"))
            {
                properties[col] = "Cancelled On";
                col++;
            }

            if (columns.Contains("CreditedOn"))
            {
                properties[col] = "Credited On";
                col++;
            }

            if (columns.Contains("PackagedOn"))
            {
                properties[col] = "PackagedOn";
                col++;
            }

            if (columns.Contains("Customer"))
            {
                properties[col] = "Customer";
                col++;
            }

            if (columns.Contains("FulfillmentStatus"))
            {
                properties[col] = "Fulfillment Status";
                col++;
            }

            if (columns.Contains("ETA Date"))
            {
                properties[col] = "ETA Date";
                col++;
            }

            if (columns.Contains("TrackingNumber"))
            {
                properties[col] = "Tracking #";
                col++;
            }
            if (columns.Contains("CreditStatus"))
            {
                properties[col] = "Credit Status";
            }

            //init list check col
            Dictionary<int, int> listCheckCol = new Dictionary<int, int>();

            for (int i = 0; i < properties.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;

                //add value for list check col: key = coln, value = maxLength
                listCheckCol.Add(i + 1, 0);

                AutoFitColumnExportExcel(worksheet, row, i + 1, listCheckCol);
            }
            worksheet.View.FreezePanes(2, 1);
            row++;
            #endregion Bind Title

            #region Bind Data
            foreach (var o in orders)
            {
                col = 1;

                if (columns.Contains("Id"))
                {
                    worksheet.Cells[row, col].Value = o.Id;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("PartnerId"))
                {
                    worksheet.Cells[row, col].Value = o.Partner;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("PartnerOrderId"))
                {
                    worksheet.Cells[row, col].Value = o.PartnerOrderId;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("Factory"))
                {
                    worksheet.Cells[row, col].Value = o.FactoryName;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("OrderType"))
                {
                    worksheet.Cells[row, col].Value = o.OrderType;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("Priority"))
                {
                    worksheet.Cells[row, col].Value = o.IsPriority ? "Y" : "N";
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("Items"))
                {
                    worksheet.Cells[row, col].Value = o.Units;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("ReceivedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.ReceivedOn.HasValue
                        ? string.Empty
                        : o.ReceivedOn.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("SlaDate"))
                {
                    worksheet.Cells[row, col].Value = !o.SlaDateOnUtc.HasValue
                        ? string.Empty
                        : o.SlaDateOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("DroppedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.DroppedOnUtc.HasValue
                        ? string.Empty
                        : o.DroppedOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("MergeBinId"))
                {
                    worksheet.Cells[row, col].Value = !string.IsNullOrEmpty(o.MergeBinId) ? o.MergeBinId : string.Empty;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("MergedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.MergedOnUtc.HasValue
                        ? string.Empty
                        : o.MergedOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("PrintedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.PrintedOnUtc.HasValue
                        ? string.Empty
                        : o.PrintedOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("ShippedOn"))
                {
                    // TDP-1585: TDP portal should not show Shipped On time for Packaged orders
                    worksheet.Cells[row, col].Value = !o.SortedOn.HasValue
                        ? string.Empty
                        : o.SortedOn.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("SortedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.SortedOn.HasValue
                        ? string.Empty
                        : o.SortedOn.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("PackagedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.ShippedOn.HasValue
                        ? string.Empty
                        : o.ShippedOn.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("CancelledOn"))
                {
                    worksheet.Cells[row, col].Value = !o.CancelledOnUtc.HasValue
                        ? string.Empty
                        : o.CancelledOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("CreditedOn"))
                {
                    worksheet.Cells[row, col].Value = !o.CreditedOnUtc.HasValue
                        ? string.Empty
                        : o.CreditedOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }
                if (columns.Contains("Customer"))
                {
                    worksheet.Cells[row, col].Value = o.Customer;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("FulfillmentStatus"))
                {
                    worksheet.Cells[row, col].Value = o.FulfillmentStatus;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("ETA Date"))
                {
                    worksheet.Cells[row, col].Value = o.EtaDateOnEst;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }

                if (columns.Contains("TrackingNumber"))
                {
                    worksheet.Cells[row, col].Value = o.TrackingNumber;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }
                if (columns.Contains("ShippingState"))
                {
                    worksheet.Cells[row, col].Value = o.ShippingState;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }
                if (columns.Contains("ShippingCountry"))
                {
                    worksheet.Cells[row, col].Value = o.ShippingCountry;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }
                if (columns.Contains("ShippingCarrier"))
                {
                    worksheet.Cells[row, col].Value = o.ShippingCarrier;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                    col++;
                }
                if (columns.Contains("CreditStatus"))
                {
                    worksheet.Cells[row, col].Value = o.CreditStatus;
                    worksheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    AutoFitColumnExportExcel(worksheet, row, col, listCheckCol);
                }

                row++;
            }

            worksheet.Cells.AutoFitColumns(0);
            #endregion Bind Data
        }

        private static void BindOrderDetailSheet(
            ExcelPackage xlPackage,
            List<ExportOrderDetailSummaryModel> details
        )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Detail");
            var row = 1;

            #region Bind Title
            var properties = new List<string>
            {
                "Type",
                "Order #",
                "Partner Order ID",
                "TSC Order ID",
                "Factory Name",
                "Received On",
                "Cancelled On",
                "Shipped on",
                "Iten Name",
                "SKU",
                "Garment ID",
                "Vendor Style",
                "Color",
                "Size",
                "Quantity",
                "Hit Count",
                "Inv Unit Price",
                "Inv Total"
            };
            for (int i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[row, i + 1].Value = properties[i];
                worksheet.Cells[row, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[row, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[row, i + 1].Style.Font.Bold = true;
                worksheet.Column(i + 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            }

            row++;
            worksheet.View.FreezePanes(2, 1);
            #endregion Bind Title

            #region Bind Data
            foreach (var detail in details)
            {
                int col = 1;

                worksheet.Cells[row, col].Value = detail.Type;
                col++;

                worksheet.Cells[row, col].Value = detail.OrderId;
                col++;

                worksheet.Cells[row, col].Value = detail.Xid;
                col++;
                
                worksheet.Cells[row, col].Value = detail.FactoryName;
                col++;

                worksheet.Cells[row, col].Value = detail.ReceivedOnUtc.ToString("MM-dd-yyyy hh:mm:ss tt");
                col++;

                worksheet.Cells[row, col].Value = !detail.CancelledOnUtc.HasValue ? string.Empty : detail.CancelledOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                col++;

                worksheet.Cells[row, col].Value = !detail.ShippedOnUtc.HasValue ? string.Empty : detail.ShippedOnUtc.Value.ToString("MM-dd-yyyy hh:mm:ss tt");
                col++;

                worksheet.Cells[row, col].Value = detail.ItemName;
                col++;

                worksheet.Cells[row, col].Value = detail.Sku;
                col++;

                worksheet.Cells[row, col].Value = detail.GarmentId;
                col++;

                worksheet.Cells[row, col].Value = detail.VendorStyle;
                col++;

                worksheet.Cells[row, col].Value = detail.Color;
                col++;

                worksheet.Cells[row, col].Value = detail.Size;
                col++;

                worksheet.Cells[row, col].Value = detail.Quantity;
                col++;

                worksheet.Cells[row, col].Value = detail.HitCount;
                col++;

                worksheet.Cells[row, col].Value = detail.InvoiceUnitPrice;
                col++;

                worksheet.Cells[row, col].Value = detail.InvoiceTotal;
                row++;
            }
            #endregion Bind Data

            worksheet.Cells.AutoFitColumns(0);
        }

        public async Task<byte[]> ExportExcelOrderAsync(
            EOrderStatusExtend orderStatus, 
            List<ExportOrderSummaryModel> orders, 
            List<ExportOrderDetailSummaryModel> details
            )
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                var orderColumns = GetColumnsOrderByOrderStatus(orderStatus);
                BindOrderSheet(xlPackage, orderColumns, orders);

                BindOrderDetailSheet(xlPackage, details);

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
