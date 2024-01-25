using System.Diagnostics;
using ExportConsoleApp.Services;
using OfficeOpenXml;

var service = new ExportRoyaltyReportService();

////Summary
////var data = service.BindSummaryData();
////var bytes = await service.ExportWeeklySummaryReconciliationAsync(data);

////Factory
//var data = service.BindData();
//var bytes = await service.ExportRoyaltyReportAsync(data,true);

//string path = Path.Combine(@"C:\Users\nguye\Downloads\Console-App-Files", Guid.NewGuid() + ".xlsx");
//File.WriteAllBytes(path, bytes);

//Process p = new Process();
//p.StartInfo.FileName = path;
//p.StartInfo.UseShellExecute = true;
//p.Start();

//p.WaitForExit();


// path to your excel file
string path = "C:\\Users\\ThuHuynh.AzureAD\\Downloads\\test intro foreword.xlsx";
FileInfo fileInfo = new FileInfo(path);
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using (var pck = new ExcelPackage(fileInfo))
{
    var ws = pck.Workbook.Worksheets.FirstOrDefault();
    var cols = ws.Dimension.Columns;
    var index = 1;
    for (int row = ws.Dimension.End.Row - ws.Dimension.Start.Row + 1; row > 1; row--)
    {
        var isEmptyRow = !(ws.Cells[row, 1, row, cols].ToList()).Any(x => x.Value != null);
        if (isEmptyRow)
        {
            continue;
        }

        index = row;
        break;
    }
}

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();


