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
string path = "C:\\Users\\ThuHuynh.AzureAD\\Downloads\\Thu-Test-Copy1-20240118-0649.xlsx";
FileInfo fileInfo = new FileInfo(path);

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
ExcelPackage package = new ExcelPackage(fileInfo);
var worksheets = package.Workbook.Worksheets.ToList();

// get number of rows and columns in the sheet

 int index = 0;
 foreach (var ws in worksheets)
 {
     int rows = ws.Dimension.Rows; // 20
    var ws2 = package.Workbook.Worksheets.Add($"Copy-{index}",
        ws);
    ws2.Cells[rows + 1, 1].Value = "AAAA";
    index++;
 }

 package.Save();

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();



