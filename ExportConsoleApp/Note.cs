using iText.Commons.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportConsoleApp
{
    public class Note
    {
        private readonly List<string> _properties = new List<string> {
            "Image Type",
            "Official Art File",
            "Virtual Catalog",
            "Wholesale",
            "SMPL Required",
            "SMPL Approved",
            "Product Approval",
            "Design Name",
            "Tags",
            "PartnerId",
            "Licensor",
            "License",
            "Sub-License",
            "Artist",
            "Program",
            "Royalty Code",
            "Royalty Amount",
            "Art ID",
            "Master Art ID",
            "Partner Art ID",
            "Product Id",
            "Style",
            "Partner SKU",
            "UPC",
            "Invoice Price",
            "Retail Price",
            "MD SKU",
            "Description",
            "Color Code",
            "Color Desc",
            "Is Light?",
            "Size Code",
            "Size Desc",
            "Size Class Code",
            "Size Class Desc",
            "Weight (LBS)",
            "Digital Print Location 1",
            "Digital Print Image URL 1",
            "Digital Preview Image URL 1",
            "Digital Print Location 2",
            "Digital Print Image URL 2",
            "Digital Preview Image URL 2",
            "Digital Print Location 3",
            "Digital Print Image URL 3",
            "Digital Preview Image URL 3",
            "Digital Print Location 4",
            "Digital Print Image URL 4",
            "Digital Preview Image URL 4",
            "Neck Label Preview Image URL",
            "Neck Label Bin Id",
            "Hang Tag Preview Image URL",
            "Hang Tag Bin Id",
            "Sticker Preview Image URL",
            "Sticker Bin Id",
            "Patch Preview Image URL",
            "Patch Bin Id"
        };
        public async Task ImportArtworkAsync()
        {
            var models = new List<string>();

            #region VALIDATE IMPORT ROWS

            Stream stream = new MemoryStream();
            using (var pck = new ExcelPackage(stream))
            {
                var ws = pck.Workbook.Worksheets.FirstOrDefault();
                for (int row = ws.Dimension.End.Row - ws.Dimension.Start.Row + 1; row > 1; row--)
                {
                    var isEmptyRow = !(ws.Cells[row, 1, row, _properties.Count].ToList()).Any(x => x.Value != null);
                    if (isEmptyRow)
                    {
                        continue;
                    }
                    var model = "aa";
                    //Log error when value invalid
                    models.Add(model);
                }
            }
            #endregion VALIDATE IMPORT ROWS

        }
    }
}

