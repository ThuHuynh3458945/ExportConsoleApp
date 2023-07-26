using System.Diagnostics;
using ExportConsoleApp.Enums;
using ExportConsoleApp.Services;

var service = new ExportOrderService();

var orderStatus = EOrderStatusExtend.InShortage;
var orders = service.BindDataOrder();
var orderDetails = service.BindDataOrderDetail();

var bytes = await service.ExportExcelOrderAsync(orderStatus, orders, orderDetails);

string path = Path.Combine(@"C:\Users\tuan\Downloads\Console-App-Files", Guid.NewGuid() + ".xlsx");
File.WriteAllBytes(path, bytes);

Process p = new Process();
p.StartInfo.FileName = path;
p.StartInfo.UseShellExecute = true;
p.Start();

p.WaitForExit();