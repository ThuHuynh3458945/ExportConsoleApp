using System.Diagnostics;
using ExportConsoleApp.Services;

var service = new ExportRoyaltyReportService();

//Summary
//var data = service.BindSummaryData();
//var bytes = await service.ExportWeeklySummaryReconciliationAsync(data);

//Factory
var data = service.BindData();
var bytes = await service.ExportRoyaltyReportAsync(data,true);

string path = Path.Combine(@"C:\Users\nguye\Downloads\Console-App-Files", Guid.NewGuid() + ".xlsx");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();