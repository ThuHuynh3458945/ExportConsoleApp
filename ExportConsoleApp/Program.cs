using System.Diagnostics;
using ExportConsoleApp.Services;

var service = new InvoiceService();
var data = service.BuildData();
var bytes = await service.ExportInvoiceAsync(data);
string path = Path.Combine(@"C:\Users\nguye\Downloads\Console-App-Files", Guid.NewGuid().ToString() + ".pdf");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();