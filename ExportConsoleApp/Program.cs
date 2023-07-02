using ExportConsoleApp;
using ExportConsoleApp.Models;
using System.Diagnostics;

var service = new ExportFulfillmentCustomerReportCardService();
var data = service.BuildData();
var bytes = service.ExportFulfillmentCustomerReportCard("Monster Digital Fulfillment - Customer Report Card (3/04/23)", data);
string path = Path.Combine(@"C:\Users\nguye\Downloads\Console-App-Files", Guid.NewGuid().ToString() + ".pdf");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();