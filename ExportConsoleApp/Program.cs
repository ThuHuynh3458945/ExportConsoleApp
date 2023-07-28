using System.Diagnostics;
using ExportConsoleApp.Services;

var service = new ExportWeeklyReconciliationService();

//Summary
//var data = service.BindSummaryData();
//var bytes = await service.ExportWeeklySummaryReconciliationAsync(data);

//Factory
var data = service.BindFactoryData();
var bytes = await service.ExportWeeklyFactoryReconciliationAsync(data);

string path = Path.Combine(@"C:\Users\tuan\Downloads\Console-App-Files", Guid.NewGuid() + ".xlsx");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();