using System;
using System.Collections;
using System.Linq;
using GemBox.Spreadsheet;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

IWebDriver driver = new ChromeDriver();

var config = JsonConvert.DeserializeObject<Root>(File.ReadAllText(@"./config.json"));
var dataConfig = JsonConvert.DeserializeObject<DataConfig.Root>(File.ReadAllText(@"./dataConfig.json"));
SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

if (config == null || dataConfig == null)
{
    Console.WriteLine("Loi: Khong doc duoc tep cau hinh");
    Console.ReadLine();
    return;
}

driver.Navigate().GoToUrl("https://ktat.cpc.vn/Hethong/Dangnhap.aspx");
Task.Delay(2000).Wait();
var usernameElement = driver.FindElement(By.Id(config.username.id));
usernameElement.Clear();
usernameElement.SendKeys(config.username.user);
var passwordElement = driver.FindElement(By.Id(config.password.id));
passwordElement.Clear();
passwordElement.SendKeys(config.password.pass);
var loginElement = driver.FindElement(By.Id(config.login.id));
loginElement.Click();
driver.Navigate().GoToUrl("https://ktat.cpc.vn/Hethong/tbat_xemDS_TB.aspx");
Task.Delay(2000).Wait();
var unitElement = driver.FindElement(By.Id(config.unit.id));
if (config.unit.includeAll == "yes")
    unitElement.Click();
var statusElement = new SelectElement(driver.FindElement(By.Id(config.status.id)));
statusElement.SelectByIndex(int.Parse(config.status.choose));
var getAllDevice = driver.FindElement(By.Id(config.getAllDevice.id));
getAllDevice.Click();
Console.WriteLine("Cho lay du lieu thiet bi");
Task.Delay(2000).Wait();
var exportXLSX = driver.FindElement(By.Id(config.exportXLSX.id));
exportXLSX.Click();
Console.WriteLine("Cho tai tep ve");
Task.Delay(7000).Wait();
var downloadFolder = config.download.folderDir;
if (Directory.Exists(downloadFolder))
{
    DirectoryInfo dirInfo = new DirectoryInfo(downloadFolder);
    var ext =  "*.xlsx";
    List<FileInfo> files = new List<FileInfo>();
    files.AddRange(dirInfo.GetFiles(ext).OrderByDescending(x => x.CreationTime).ToArray());
    if (files == null || files.Count == 0){
        Console.WriteLine("Loi! Khong tim thay tep da tai ve hoac het thoi gian cho nhung tep van chua tai xong");
        Console.ReadLine();
        return;
    }
    var lastestFile = (FileInfo)files[0];
    if (DateTime.Now.Subtract(lastestFile.CreationTime) >= TimeSpan.FromDays(1))
    {
        Console.WriteLine("Loi! Tep da tai ve khong phai la tep moi nhat trong ngay");
        Console.ReadLine();
        return;
    }
    var workBook = ExcelFile.Load(lastestFile.FullName);
    var workSheet = workBook.Worksheets[0];
    List<Device> deviceList = new List<Device>();
    for (int row = 1; row < workSheet.Rows.Count; row++)
    {
        deviceList.Add(new Device()
        {
            Id = workSheet.Cells[row, dataConfig.idIndexCol].Value?.ToString(),
            DeviceName = (workSheet.Cells[row, dataConfig.deviceNameIndexCol]).Value?.ToString(),
            UnitName = (workSheet.Cells[row, dataConfig.unitNameIndexCol])?.Value?.ToString(),
            LastestUsedDate =(workSheet.Cells[row, dataConfig.lastestUsedDateIndexCol])?.Value?.ToString(), 
            CustomerName = workSheet.Cells[row, dataConfig.customerNameIndexCol]?.Value?.ToString(), 
            DeviceManufacturerName =(workSheet.Cells[row, dataConfig.deviceManufacturerNameIndexCol])?.Value?.ToString(), 
            DeviceManufactureCountry =(workSheet.Cells[row, dataConfig.deviceManufactureCountryIndexCol])?.Value?.ToString(),
            DeviceYearOf =(workSheet.Cells[row, dataConfig.deviceYearOfIndexCol])?.Value?.ToString(),
            DeviceStatus =(workSheet.Cells[row, dataConfig.deviceStatusIndexCol])?.Value?.ToString(),
            DeviceProductionCode =(workSheet.Cells[row, dataConfig.deviceProductionCodeIndexCol])?.Value?.ToString()
        });
    }
    Console.WriteLine("Tong so thiet bi duoc thu thap la: " + deviceList.Count.ToString());
}
else
{
    Console.WriteLine("Loi: Khong tim thay thu muc da duoc chi dinh");
    Console.ReadLine();
}
Console.WriteLine("Nhan Enter de thoat!");
Console.ReadLine();
//driver.Quit();
