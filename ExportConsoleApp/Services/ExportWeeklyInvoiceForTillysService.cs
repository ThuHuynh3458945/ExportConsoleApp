using ExportConsoleApp.Models;
using System.Text;
using ExportConsoleApp.Helpers;

namespace ExportConsoleApp.Services
{
    public class ExportWeeklyInvoiceForTillysService
    {
        public List<InvoiceTillysModel> BuildData()
        {
            return new List<InvoiceTillysModel>
            {
                new InvoiceTillysModel
                {
                    InvoiceWeek = "WK24-23",
                    PartnerId = "tillys",
                    Factory = "Juarez",
                    OrderDate = "6/8/2023 19:09",
                    ShipDate = "6/11/2023 10:13",
                    CancelDate = "",
                    CreditedDate = "",
                    OrderId = "23949617",
                    PartnerOrderId = "12646003-2",
                    MiscOrderId = "1063029",
                    Upc = "",
                    Sku = "002-002-002",
                    PartnerSku = "467213002",
                    PartnerBlankSku = "",
                    License = "",
                    PartnerItem = "Hippie Skeleton White T Shirt",
                    StyleDescription = "Hippie Skeleton White T Shirt",
                    SizeClass = "Mens",
                    Size = "Medium",
                    Color = "White",
                    Front = "1",
                    Back = "",
                    Left = "",
                    Right = "",
                    Quantity = "1",
                    FulfillmentUnitCost = "0",
                    GarmentUnitCost = "9.68",
                    LineTotal = "9.68",
                    PackagingCost = "0",
                    ShipCost = "0",
                    CreditAmount = "0",
                    FulfillmentType = "Digital Print",
                    Customer = "Miguel Gallardo",
                    AddressLine1 = "1056 CHURCH ST",
                    AddressLine2 = "",
                    City = "BEAUMONT",
                    State = "TX",
                    ZipCode = "77705-2202",
                    Country = "US",
                    ShippingCarrier = "FDX",
                    ShippingPriority = "2010",
                    TrackingNumber = "74890308207020013122"
                }
            };
        }

        public async Task<byte[]> ExportCsvWeeklyInvoiceForTillysAsync(List<InvoiceTillysModel> items)
        {
            #region Title
            var headerList = new List<string>
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
                    "TrackingNumber"
                };

            var str = new StringBuilder();
            foreach (var header in headerList)
            {
                str.Append(header + ",");
            }
            str.Append("\r\n");
            #endregion Title

            #region Bind Data Detail

            foreach (var item in items)
            {
                str.Append(item.InvoiceWeek.RemoveComma() + ",");
                str.Append(item.PartnerId.RemoveComma() + ",");
                str.Append(item.Factory.RemoveComma() + ",");
                str.Append(item.OrderDate.RemoveComma() + ",");
                str.Append(item.ShipDate.RemoveComma() + ",");
                str.Append(item.CancelDate.RemoveComma() + ",");
                str.Append(item.CreditedDate.RemoveComma() + ",");
                str.Append(item.OrderId.RemoveComma() + ",");
                str.Append(item.PartnerOrderId.RemoveComma() + ",");
                str.Append(item.MiscOrderId.RemoveComma() + ",");
                str.Append(item.Upc.RemoveComma() + ",");
                str.Append(item.Sku.RemoveComma() + ",");
                str.Append(item.PartnerSku.RemoveComma() + ",");
                str.Append(item.PartnerBlankSku.RemoveComma() + ",");
                str.Append(item.License.RemoveComma() + ",");
                str.Append(item.PartnerItem.RemoveComma() + ",");
                str.Append(item.StyleDescription.RemoveComma() + ",");
                str.Append(item.SizeClass.RemoveComma() + ",");
                str.Append(item.Size.RemoveComma() + ",");
                str.Append(item.Color.RemoveComma() + ",");
                str.Append(item.Front.RemoveComma() + ",");
                str.Append(item.Back.RemoveComma() + ",");
                str.Append(item.Left.RemoveComma() + ",");
                str.Append(item.Right.RemoveComma() + ",");
                str.Append(item.Quantity.RemoveComma() + ",");
                str.Append(item.FulfillmentUnitCost.RemoveComma() + ",");
                str.Append(item.GarmentUnitCost.RemoveComma() + ",");
                str.Append(item.LineTotal.RemoveComma() + ",");
                str.Append(item.PackagingCost.RemoveComma() + ",");
                str.Append(item.ShipCost.RemoveComma() + ",");
                str.Append(item.CreditAmount.RemoveComma() + ",");
                str.Append(item.FulfillmentType.RemoveComma() + ",");
                str.Append(item.Customer.RemoveComma() + ",");
                str.Append(item.AddressLine1.RemoveComma() + ",");
                str.Append(item.AddressLine2.RemoveComma() + ",");
                str.Append(item.City.RemoveComma() + ",");
                str.Append(item.State.RemoveComma() + ",");
                str.Append(item.ZipCode.RemoveComma() + ",");
                str.Append(item.Country.RemoveComma() + ",");
                str.Append(item.ShippingCarrier.RemoveComma() + ",");
                str.Append(item.ShippingPriority.RemoveComma() + ",");
                str.Append("'" + item.TrackingNumber.RemoveComma() + ",");
                str.Append("\r\n");
            }
            #endregion Bind Data Detail

            return Encoding.UTF8.GetBytes(str.ToString());
        }
    }
}
