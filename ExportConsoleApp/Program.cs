//using System.Diagnostics;
//using ExportConsoleApp.Services;
//using OfficeOpenXml;

//var service = new ExportRoyaltyReportService();

//////Summary
//////var data = service.BindSummaryData();
//////var bytes = await service.ExportWeeklySummaryReconciliationAsync(data);

//////Factory
////var data = service.BindData();
////var bytes = await service.ExportRoyaltyReportAsync(data,true);

////string path = Path.Combine(@"C:\Users\nguye\Downloads\Console-App-Files", Guid.NewGuid() + ".xlsx");
////File.WriteAllBytes(path, bytes);

////Process p = new Process();
////p.StartInfo.FileName = path;
////p.StartInfo.UseShellExecute = true;
////p.Start();

////p.WaitForExit();

////await Task.WhenAll(service.Method1(), service.Method2());

////// path to your excel file
////string path = "C:\\Users\\ThuHuynh.AzureAD\\Downloads\\test intro foreword.xlsx";
////FileInfo fileInfo = new FileInfo(path);
////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
////using (var pck = new ExcelPackage(fileInfo))
////{
////    var ws = pck.Workbook.Worksheets.FirstOrDefault();
////    var cols = ws.Dimension.Columns;
////    var index = 1;
////    for (int row = ws.Dimension.End.Row - ws.Dimension.Start.Row + 1; row > 1; row--)
////    {
////        var isEmptyRow = !(ws.Cells[row, 1, row, cols].ToList()).Any(x => x.Value != null);
////        if (isEmptyRow)
////        {
////            continue;
////        }

////        index = row;
////        break;
////    }
////}

////Process p = new Process();
////p.StartInfo.FileName = path;
////p.StartInfo.UseShellExecute = true;
////p.Start();

////p.WaitForExit();

//var method1Task = service.Method1();
//var method2Task = service.Method2();

//await Task.WhenAll(method1Task, method2Task);

//var b = method1Task.Result;
//var c = method2Task.Result;


using ExportConsoleApp;

namespace NFQ_Kata
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("OMGHAI!");
            List<Item> Items = new List<Item>{
            new Item {Name = "+5 Dexterity Vest", SellIn = 10, Quality = 20},
            new Item {Name = "Aged Brie", SellIn = 2, Quality = 0},
            new Item {Name = "Elixir of the Mongoose", SellIn = 5, Quality = 7},
            new Item {Name = "Sulfuras, Hand of Ragnaros", SellIn = 0, Quality = 80},
            new Item {Name = "Sulfuras, Hand of Ragnaros", SellIn = -1, Quality = 80},
            new Item {Name = "Backstage passes to a TAFKAL80ETC concert",SellIn = 15,Quality = 20},
            new Item {Name = "Backstage passes to a TAFKAL80ETC concert",SellIn = 10,Quality = 49},
            new Item {Name = "Backstage passes to a TAFKAL80ETC concert",SellIn = 5,Quality = 49},
            // this conjured item does not work properly yet
            new Item {Name = "Conjured Mana Cake", SellIn = 3, Quality = 6}
        };
            var app = new GildedRose(Items);
            for (var i = 0; i < 31; i++)
            {
                Console.WriteLine("-------- day " + i + " --------");
                Console.WriteLine("name, sellIn, quality");

                for (var j = 0; j < Items.Count; j++)
                {
                    Console.WriteLine($"{Items[j].Name}, {Items[j].SellIn}, {Items[j].Quality}");
                }
                Console.WriteLine("");
                //app.UpdateQuality1();
                app.UpdateQuality();
            }
        }
    }
}