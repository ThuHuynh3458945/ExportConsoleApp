using System.Diagnostics;
using ExportConsoleApp.Services;

var service = new ExportWeeklyInvoiceForTillysService();
var data = service.BuildData();
var bytes = await service.ExportCsvWeeklyInvoiceForTillysAsync(data);
string path = Path.Combine(@"C:\Users\tuan\Downloads\Console-App-Files", Guid.NewGuid() + ".csv");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();