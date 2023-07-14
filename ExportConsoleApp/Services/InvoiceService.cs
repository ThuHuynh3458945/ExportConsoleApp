using ExportConsoleApp.Heplers;
using ExportConsoleApp.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExportConsoleApp.Services
{
    public class InvoiceService
    {
        public InvoiceForReportModel BuildData()
        {
            return new InvoiceForReportModel
            {
                PartnerId = "teefuryvip",
                FromDate = DateTime.UtcNow,
                EndDate = DateTime.UtcNow,

                MdInvoice = new MdInvoiceForReportModel
                {
                    Total = new InvoiceTotalForReportModel
                    {
                        TotalOrders = 100,
                        TotalQuantity = 100,
                        TotalLine = 100,
                        BagFee = 100,
                        PackagingFee = 100,
                        ShippingFee = 100,
                        ShipCost = 100,
                        TotalCreditAmount = 100
                    },
                    Details = new List<InvoiceDetailForReportModel>
                    {
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "teefuryvip",
                            FactoryName = "Miami",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "1",
                            MiscOrderId = "12",
                            TypeOrder = "SAMPLE",
                            Upc = "123",
                            Sku = "1231",
                            PartnerSku = "4",
                            PartnerBlankSku = "5",
                            License = "6",
                            PartnerItem = "7",
                            Description = "6",
                            SizeClassDesc = "8",
                            SizeDesc = "9",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "1",
                            Customer = "2",
                            Address1 = "34",
                            Address2 = "4",
                            City = "5",
                            State = "6",
                            ZipCode = "7",
                            Country = "8",
                            ShippingCarrier = "1",
                            ShippingPriority = "2",
                            TrackingNumber = "3",
                            Notes = "4"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        }
                    }
                },
                TscInvoice = new MdInvoiceForReportModel
                {
                    Total = new InvoiceTotalForReportModel
                    {
                        TotalOrders = 100,
                        TotalQuantity = 100,
                        TotalLine = 100,
                        BagFee = 100,
                        PackagingFee = 100,
                        ShippingFee = 100,
                        ShipCost = 100,
                        TotalCreditAmount = 100
                    },
                    Details = new List<InvoiceDetailForReportModel>
                    {
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        },
                        new InvoiceDetailForReportModel
                        {
                            InvoiceWeek = "123123123",
                            PartnerId = "123123123",
                            FactoryName = "123123123",
                            OrderDate = DateTime.UtcNow,
                            ShipDate = DateTime.UtcNow,
                            CancelledOnUtc = DateTime.UtcNow,
                            CreditedOnUtc = DateTime.UtcNow,
                            OrderId = 1000,
                            PartnerOrderId = "123123123",
                            MiscOrderId = "123123123",
                            TypeOrder = "SAMPLE",
                            Upc = "123123123",
                            Sku = "123123123",
                            PartnerSku = "123123123",
                            PartnerBlankSku = "123123123",
                            License = "123123123",
                            PartnerItem = "123123123",
                            Description = "123123123",
                            SizeClassDesc = "123123123",
                            SizeDesc = "123123123",
                            ColorDesc = "123123123",
                            Front = 1,
                            Back = 2,
                            Left = 3,
                            Right = 4,
                            Quantity = 1000,
                            FulfillmentUnitCost = 1000,
                            GarmentUnitCost =1000,
                            TotalLine = 1000,
                            PackagingCost = 1000,
                            ShipCost = 1000,
                            Type = "123123123",
                            Customer = "123123123",
                            Address1 = "123123123",
                            Address2 = "123123123",
                            City = "123123123",
                            State = "123123123",
                            ZipCode = "123123123",
                            Country = "123123123",
                            ShippingCarrier = "123123123",
                            ShippingPriority = "123123123",
                            TrackingNumber = "123123123",
                            Notes = "123123123"
                        }
                    }
                },
                OtsInvoices = new List<TscInvoiceDetailForReportModel>
                {
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    },
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    },
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    },
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    },
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    },
                    new TscInvoiceDetailForReportModel
                    {
                        MachineVendor = "TSC",
                        Company = "TSC123",
                        PoNumber = "123123",
                        CancelDate = DateTime.UtcNow,
                        Style = "test",
                        Description = "test",
                        Color = "white",
                        PrintDate = DateTime.UtcNow,
                        PrintQty = 1000,
                        UnitCost = 10,
                        InvoicePrice = 100
                    }
                },
                Credits = new List<AuthorizedCreditForReportModel>
                {
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    },
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    },
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    },
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    },
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    },
                    new AuthorizedCreditForReportModel
                    {
                        InvoiceWeek = "WK-25",
                        PartnerId = "tsc",
                        CreditedOnUtc = DateTime.UtcNow,
                        OrderId = 123123,
                        PartnerOrderId = "123123123",
                        MiscOrderId = "123123123123",
                        CreditAmount = 100,
                        CreditReason = "Tach"
                    }
                }
            };
        }

        private static void BindEmptySheet(
            ExcelPackage xlPackage,
            string sheetName,
            DateTime fromDate,
            DateTime endDate)
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add(sheetName);

            const int row = 2;
            const int col = 1;

            worksheet.Cells[row, col].Value = $"No orders were shipped from {fromDate:MM/dd/yyyy} through {endDate:MM/dd/yyyy}";
            worksheet.Cells[row, col].Style.Font.Size = 16;
            worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, col].Style.Font.Bold = true;
            worksheet.Cells.AutoFitColumns();
        }

        private static void BindInvoicesSheet(
            ExcelPackage xlPackage,
            string sheetName,
            MdInvoiceForReportModel data
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add(sheetName);
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
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }

            row++;
            #endregion Title

            #region Bind Data Detail

            foreach (var invoice in data.Details)
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

            var total = data.Total;
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

            worksheet.Cells[row, 2].Value = total.TotalOrders;
            row++;

            //Total Units
            worksheet.Cells[row, 1].Value = "Total Units";
            worksheet.Cells[row, 2].Value = total.TotalQuantity;
            row++;

            //Total line
            worksheet.Cells[row, 1].Value = "Line Total";
            worksheet.Cells[row, 2].Value = total.TotalLine;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            //Bag Fee
            worksheet.Cells[row, 1].Value = "Bag Fee";
            worksheet.Cells[row, 2].Value = total.BagFee;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            worksheet.Cells[row, 1].Value = "Packaging Fee";
            worksheet.Cells[row, 2].Value = total.PackagingFee;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            worksheet.Cells[row, 1].Value = "Shipping Fee";
            worksheet.Cells[row, 2].Value = total.ShippingFee;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            worksheet.Cells[row, 1].Value = "Credit Total";
            worksheet.Cells[row, 2].Value = total.TotalCreditAmount;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";
            row++;

            //Invoice Total                    
            worksheet.Cells[row, 1].Value = "Invoice Total";
            worksheet.Cells[row, 2].Value = total.TotalLine + total.ShippingFee + total.PackagingFee + total.BagFee + total.TotalCreditAmount;
            worksheet.Cells[row, 2].Style.Numberformat.Format = "0.00";

            worksheet.Cells[row - 7, 1, row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row - 7, 1, row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row - 7, 1, row, 1].Style.Font.Bold = true;

            worksheet.Cells.AutoFitColumns();
            #endregion Bind Data Total

            worksheet.View.FreezePanes(2, 1);
        }

        private static void BindOtsInvoicesSheet(
            ExcelPackage xlPackage,
            string sheetName,
            List<TscInvoiceDetailForReportModel> invoices
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add(sheetName);
            var row = 1;

            #region Title
            var properties = new List<string>
                {
                    "Machine/Vendor",
                    "Company",
                    "PO #",
                    "Cancel Date",
                    "Style",
                    "Description",
                    "Color",
                    "Print Date",
                    "Print Qty",
                    "Unit Cost",
                    "Invoice Price"
                };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
            row++;
            #endregion Title

            #region Bind Data Detail
            foreach (var item in invoices)
            {
                var col = 1;

                //Machine/Vendor
                worksheet.Cells[row, col].Value = item.MachineVendor;
                col++;

                //Company
                worksheet.Cells[row, col].Value = item.Company;
                col++;

                //PO #
                worksheet.Cells[row, col].Value = item.PoNumber;
                col++;

                //Cancel Date
                worksheet.Cells[row, col].Value = DateTimeHelper.ConvertToEstTime(item.CancelDate).ToString("M/d/yyyy");
                col++;

                //Style
                worksheet.Cells[row, col].Value = item.Style;
                col++;

                //Description
                worksheet.Cells[row, col].Value = item.Description;
                col++;

                //Color
                worksheet.Cells[row, col].Value = item.Color;
                col++;

                //Print Date
                worksheet.Cells[row, col].Value = DateTimeHelper.ConvertToEstTime(item.PrintDate).ToString("M/d/yyyy");
                col++;

                //Print Qty
                worksheet.Cells[row, col].Value = item.PrintQty;
                col++;

                //Unit Cost
                worksheet.Cells[row, col].Value = item.UnitCost;
                col++;

                //Invoice Price
                worksheet.Cells[row, col].Value = item.InvoicePrice;
                row++;

                row++;
            }
            #endregion Bind Data Detail

            #region Bind Data Total
            //Total
            worksheet.Cells[row, 9].Value = invoices.Sum(p => p.PrintQty);
            worksheet.Cells[row, 11].Value = invoices.Sum(p => p.InvoicePrice);
            row++;

            //Blank row
            worksheet.Cells[row, 1].Value = "";
            row++;

            //Total Orders
            worksheet.Cells[row, 1].Value = "Total Orders";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = invoices.GroupBy(p => p.PoNumber).Count();
            row++;

            //Total Units
            worksheet.Cells[row, 1].Value = "Total Units";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = invoices.Sum(p => p.PrintQty);
            row++;

            //Invoice Total
            worksheet.Cells[row, 1].Value = "Invoice Total";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = invoices.Sum(p => p.InvoicePrice);
            #endregion Bind Data Total

            worksheet.Cells.AutoFitColumns();
            worksheet.View.FreezePanes(2, 1);
        }

        private static void BindCreditSheet(
            ExcelPackage xlPackage,
            List<AuthorizedCreditForReportModel> credits
            )
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Credit");
            var row = 1;

            #region Title
            var properties = new List<string>
                {
                    "Credit Week",
                    "Partner Id",
                    "Credited Date",
                    "Order ID",
                    "Partner Order ID",
                    "Misc Order ID",
                    "Credit Amount",
                    "Credit Reason"
                };

            for (var i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i];
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
            row++;
            #endregion Title

            #region Bind Data Detail
            decimal totalCreditAmount = 0;
            foreach (var item in credits)
            {
                var col = 1;

                worksheet.Cells[row, col].Value = item.InvoiceWeek;
                col++;

                worksheet.Cells[row, col].Value = item.PartnerId;
                col++;

                worksheet.Cells[row, col].Value = item.CreditedOnUtc == null ? string.Empty : item.CreditedOnUtc.Value.ToString("MM/dd/yyyy");
                col++;

                worksheet.Cells[row, col].Value = item.OrderId;
                col++;

                worksheet.Cells[row, col].Value = item.PartnerOrderId;
                col++;

                worksheet.Cells[row, col].Value = item.MiscOrderId;
                col++;

                worksheet.Cells[row, col].Value = item.CreditAmount;
                worksheet.Cells[row, col].Style.Numberformat.Format = "0.00";
                col++;
                totalCreditAmount += item.CreditAmount;

                worksheet.Cells[row, col].Value = item.CreditReason;

                row++;
            }
            #endregion Bind Data Detail

            #region Bind Data Total
            worksheet.Cells[row, 7].Value = totalCreditAmount > 0 ? totalCreditAmount : string.Empty;
            worksheet.Cells[row, 7].Style.Numberformat.Format = "0.00";
            row++;

            //Total Orders
            worksheet.Cells[row, 1].Value = "Total Orders";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = credits.Count > 0 ? credits.Count : string.Empty;
            row++;

            //Total Units
            worksheet.Cells[row, 1].Value = "Credit Total";
            worksheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(184, 204, 228));
            worksheet.Cells[row, 1].Style.Font.Bold = true;
            worksheet.Cells[row, 1].AutoFitColumns();
            worksheet.Cells[row, 2].Value = totalCreditAmount > 0 ? totalCreditAmount : string.Empty;

            worksheet.Cells.AutoFitColumns(0);
            #endregion Bind Data Total

            worksheet.View.FreezePanes(2, 1);
        }

        public async Task<byte[]> ExportInvoiceAsync(InvoiceForReportModel data)
        {
            using var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var xlPackage = new ExcelPackage(stream))
            {
                if (data.MdInvoice == null && data.OtsInvoices.Any())
                {
                    BindEmptySheet(xlPackage, "Invoices", data.FromDate, data.EndDate);
                }
                else
                {
                    #region MD Invoices Sheet
                    var sheetName = (data.PartnerId == "teefuryvip" || data.PartnerId == "tsc")
                        ? "Invoices (MD)"
                        : "Invoices";

                    if (data.MdInvoice == null)
                    {
                        BindEmptySheet(xlPackage, sheetName, data.FromDate, data.EndDate);
                    }
                    else
                    {
                        BindInvoicesSheet(xlPackage, sheetName, data.MdInvoice);
                    }
                    #endregion MD Invoices Sheet

                    #region TSC Invoices Sheet
                    switch (data.PartnerId)
                    {
                        case "teefuryvip":
                            {
                                if (data.TscInvoice == null)
                                {
                                    BindEmptySheet(xlPackage, "Invoices (TSC)", data.FromDate, data.EndDate);
                                }
                                else
                                {
                                    BindInvoicesSheet(xlPackage, "Invoices (TSC)", data.TscInvoice);
                                }
                                break;
                            }
                        case "tsc":
                            {
                                if (!data.OtsInvoices.Any())
                                {
                                    BindEmptySheet(xlPackage, "Invoices (TSC)", data.FromDate, data.EndDate);
                                }
                                else
                                {
                                    BindOtsInvoicesSheet(xlPackage, "Invoices (TSC)", data.OtsInvoices);
                                }
                                break;
                            }
                    }
                    #endregion TSC Invoices Sheet

                    //Credit Sheet
                    BindCreditSheet(xlPackage, data.Credits);
                }

                await xlPackage.SaveAsync();
            }

            return stream.ToArray();
        }
    }
}
