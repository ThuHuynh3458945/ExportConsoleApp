namespace ExportConsoleApp.Models
{
    public class ExportMasterBolModel
    {
        public int FactoryId { get; set; }
        public ExportMasterBolDataModel? Summary { get; set; }
        public List<ExportMasterBolDataModel> Pallets { get; set; }

        public ExportMasterBolModel()
        {
            FactoryId = 1;
            Summary = new ExportMasterBolDataModel
            {
                SheetName = "Master BOL #",
                MovementNumber = "123",
                BillOfLading = 1,
                TotalPackages = 10,
                TotalPallets = 1,
                TotalNetWeight = 1,
                TotalGrossWeight = 1,
                TotalValue = 1,
                Details = new List<ExportMasterBolDetailInfoModel>
                {
                    new ExportMasterBolDetailInfoModel
                    {
                        PartGtin = "00052987920314",
                        UOM = "EACH",
                        QtyPcs = 5,
                        UnitCost = 5,
                        TotalValue = 5,
                        TotalNetWeight = 5,
                        TotalGrossWeight = 5,
                        CooName = "tach",
                        MexicanHtsCode = "61091002",
                        UsHts = "6109100011",
                    },
                    new ExportMasterBolDetailInfoModel
                    {
                        PartGtin = "00052987920314",
                        UOM = "EACH",
                        QtyPcs = 5,
                        UnitCost = 5,
                        TotalValue = 5,
                        TotalNetWeight = 5,
                        TotalGrossWeight = 5,
                        CooName = "tach",
                        MexicanHtsCode = "61091002",
                        UsHts = "6109100011",
                    }
                }
            };
            Pallets = new List<ExportMasterBolDataModel>
            {
                new ExportMasterBolDataModel
                {
                    SheetName = "Pallet 1",
                    MovementNumber = "123",
                    BillOfLading = 1,
                    TotalPackages = 10,
                    TotalPallets = 1,
                    TotalNetWeight = 1,
                    TotalGrossWeight = 1,
                    TotalValue = 1,
                    Details = new List<ExportMasterBolDetailInfoModel>
                    {
                        new ExportMasterBolDetailInfoModel
                        {
                            PartGtin = "00052987920314",
                            UOM = "EACH",
                            QtyPcs = 5,
                            UnitCost = 5,
                            TotalValue = 5,
                            TotalNetWeight = 5,
                            TotalGrossWeight = 5,
                            CooName = "tach",
                            MexicanHtsCode = "61091002",
                            UsHts = "6109100011",
                        },
                        new ExportMasterBolDetailInfoModel
                        {
                            PartGtin = "00052987920314",
                            UOM = "EACH",
                            QtyPcs = 5,
                            UnitCost = 5,
                            TotalValue = 5,
                            TotalNetWeight = 5,
                            TotalGrossWeight = 5,
                            CooName = "tach",
                            MexicanHtsCode = "61091002",
                            UsHts = "6109100011",
                        }
                    }
                }
            };
        }
    }

    public class ExportMasterBolDataModel
    {
        public string SheetName { get; set; } = null!;
        public string? MovementNumber { get; set; }
        public int BillOfLading { get; set; }
        public int TotalPackages { get; set; }
        public int TotalPallets { get; set; }
        public decimal TotalNetWeight { get; set; }
        public decimal TotalGrossWeight { get; set; }
        public decimal TotalValue { get; set; }

        public List<ExportMasterBolDetailInfoModel> Details { get; set; } = new List<ExportMasterBolDetailInfoModel>();
    }

    public class ExportMasterBolDetailInfoModel
    {
        public string? PartGtin { get; set; }
        ///Harcode = "EACH";
        public string? UOM { get; set; }
        public int QtyPcs { get; set; }
        public decimal UnitCost { get; set; }
        public decimal TotalValue { get; set; }
        public decimal TotalNetWeight { get; set; }
        public decimal TotalGrossWeight { get; set; }
        public string? CooName { get; set; }
        public string? MexicanHtsCode { get; set; }
        public string? UsHts { get; set; }
    }

    public enum EFactory
    {
        All = 0,
        Miami = 1,
        Juarez = 2,
        Tijuana = 3,
    }

    public class ExportSection321Model
    {
        public string? ShipmentControlNumber { get; set; }
        public string? ShipmentType { get; set; }
        public string? ShipperName { get; set; }
        public string? ShipperAdddress { get; set; }
        public string? ShipperCity { get; set; }
        public string? ShipperCountry { get; set; }
        public string? ShipperState { get; set; }
        public string? ShipperPostal { get; set; }
        public string? ShipperPortofLading { get; set; }
        public string? ConsigneeName { get; set; }
        public string? ConsigneeAddress { get; set; }
        public string? ConsigneeCity { get; set; }
        public string? ConsigneeCountry { get; set; }
        public string? ConsigneeState { get; set; }
        public string? ConsigneePostal { get; set; }
        public string? ProductDescription { get; set; }
        public int ProductQty { get; set; }
        public string? ProductUOM { get; set; }
        public decimal ProductWeight { get; set; }
        public string? ProductUnitofWeight { get; set; }
        public decimal ProductValue { get; set; }
        public string? CustReference { get; set; }
        public string? USPortArr { get; set; }
        public string? FnPortLoading { get; set; }
        public string? FnPortReceipt { get; set; }
        public string? Origin { get; set; }
    }
}
