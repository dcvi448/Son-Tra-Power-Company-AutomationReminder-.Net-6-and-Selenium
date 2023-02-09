using System;
using System.Collections;
using System.Linq;
using System.Net.Http.Json;
using System.Text;
using System.Text.RegularExpressions;
using GemBox.Spreadsheet;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

IWebDriver driver = new ChromeDriver();

var config = JsonConvert.DeserializeObject<Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/config.json"));
var dataConfig = JsonConvert.DeserializeObject<DataConfig.Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/dataConfig.json"));
SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

if (config == null || dataConfig == null)
{
    Console.WriteLine("Loi: Khong doc duoc tep cau hinh");
    Console.ReadLine();
    return;
}
#region Loading the ktat page then download a device list
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
statusElement.SelectByValue(config.status.choose);

if (config.unitName.isUse == "yes")
{
    var unitNameElement = new SelectElement(driver.FindElement(By.Id(config.unitName.id)));
    unitNameElement.SelectByValue(config.unitName.choose);
}

if (config.deviceType.isUse == "yes")
{
    var deviceTypeElement = new SelectElement(driver.FindElement(By.Id(config.deviceType.id)));
    deviceTypeElement.SelectByValue(config.deviceType.choose);
}

if (config.deviceName.isUse == "yes")
{
    var deviceNameElement = new SelectElement(driver.FindElement(By.Id(config.deviceName.id)));
    deviceNameElement.SelectByValue(config.deviceName.choose);
}

if (config.levelA.isUse == "yes")
{
    var levelAElement = new SelectElement(driver.FindElement(By.Id(config.levelA.id)));
    levelAElement.SelectByValue(config.levelA.choose);
}

if (config.team.isUse == "yes")
{
    var teamElement = new SelectElement(driver.FindElement(By.Id(config.team.id)));
    teamElement.SelectByValue(config.team.choose);
}

if (config.toDate.isUse == "yes")
{
    var toDateElement = new SelectElement(driver.FindElement(By.Id(config.toDate.id)));
    toDateElement.SelectByValue(config.toDate.choose);
}

var getAllDevice = driver.FindElement(By.Id(config.getAllDevice.id));
getAllDevice.Click();
Console.WriteLine("Cho lay du lieu thiet bi");
Task.Delay(2000).Wait();
var exportXLSX = driver.FindElement(By.Id(config.exportXLSX.id));
exportXLSX.Click();
Console.WriteLine("Cho tai tep ve");
Task.Delay(7000).Wait();
#endregion

var downloadFolder = config.download.folderDir;
if (Directory.Exists(downloadFolder))
{
    #region Check the downloaded file then read it
    DirectoryInfo dirInfo = new DirectoryInfo(downloadFolder);
    var ext = "*.xlsx";
    List<FileInfo> files = new List<FileInfo>();
    files.AddRange(dirInfo.GetFiles(ext).OrderByDescending(x => x.CreationTime).ToArray());
    if (files == null || files.Count == 0)
    {
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
            UnitName = (workSheet.Cells[row, dataConfig.unitNameIndexCol]).Value?.ToString(),
            LastestUsedDate = (workSheet.Cells[row, dataConfig.lastestUsedDateIndexCol]).Value?.ToString(),
            CustomerName = workSheet.Cells[row, dataConfig.customerNameIndexCol].Value?.ToString(),
            DeviceManufacturerName = (workSheet.Cells[row, dataConfig.deviceManufacturerNameIndexCol]).Value?.ToString(),
            DeviceManufactureCountry = (workSheet.Cells[row, dataConfig.deviceManufactureCountryIndexCol]).Value?.ToString(),
            DeviceYearOf = (workSheet.Cells[row, dataConfig.deviceYearOfIndexCol]).Value?.ToString(),
            DeviceStatus = (workSheet.Cells[row, dataConfig.deviceStatusIndexCol]).Value?.ToString(),
            DeviceProductionCode = (workSheet.Cells[row, dataConfig.deviceProductionCodeIndexCol]).Value?.ToString(),
            RemainingDay = null
        });
    }
    #endregion
    foreach (var item in deviceList)
    {
        if (item.DeviceStatus != null && item.DeviceStatus != "")
        {
            try
            {
                item.RemainingDay = Regex.Match(item.DeviceStatus, @"\d+").Value.ToString();
            }
            catch
            {
                item.RemainingDay = null;
            }
        }
    }
    Console.WriteLine("Tong so thiet bi duoc thu thap la: " + deviceList.Count.ToString());
    Console.WriteLine("Bat dau loc danh sach cac thiet bi thoa man dien kien");
    var deviceListNeedInform = deviceList.Where(item => item.RemainingDay != null).Where(item => int.Parse(item.RemainingDay) >= int.Parse(config.alert.fromDay) && int.Parse(item.RemainingDay) <= int.Parse(config.alert.toDay)).ToList();
    Console.WriteLine("Tong so thiet bi thoa man dieu kien tu ngay {0} toi ngay {1}: {2}", config.alert.fromDay, config.alert.toDay, deviceListNeedInform.Count);
    Console.WriteLine("Danh sach gom co: ");
    foreach (var item in deviceListNeedInform)
    {
        Console.WriteLine("{0}\t{1}\t{2}", item.DeviceName, item.LastestUsedDate, item.DeviceStatus);
    }
    Console.WriteLine("Gui thong bao toi Eoffice");
    #region Send device list to Eoofice
    try
    {
        using (var client = new HttpClient())
        {
            client.BaseAddress = new Uri(config.eoffice.apiUrl);
            client.DefaultRequestHeaders.Add("api-version", "2");
            client.DefaultRequestHeaders.Add("Authorization", config.eoffice.token);

            var header = config.alert.message.header.Replace("{fromDay}", config.alert.fromDay).Replace("{toDay}", config.alert.toDay)
                                                    .Replace("{sumOfDevice}", deviceListNeedInform.Count.ToString());
    
            var footer = config.alert.message.footer;
            var content = "";
            if (deviceListNeedInform.Count > 0)
            {
                foreach (var item in deviceListNeedInform)
                {
                    content += config.alert.message.content.Replace("{deviceName}", item.DeviceName).Replace("{unitName}", item.UnitName)
                                .Replace("{lastestUsedDate}", item.LastestUsedDate).Replace("{customerName}", item.CustomerName)
                                .Replace("{deviceProductionCode}", item.DeviceProductionCode).Replace("{deviceManufacturerName}", item.DeviceManufacturerName)
                                .Replace("{deviceManufactureCountry}", item.DeviceManufactureCountry).Replace("{deviceYearOf}", item.DeviceYearOf)
                                .Replace("{deviceStatus}", item.DeviceStatus);
                    ;
                }
            }

            var userNameToSend = config.alert.users;
            var dataContent = header + content + footer;

            var payload = JsonConvert.SerializeObject(new
            {
                usernames = userNameToSend,
                noi_dung = dataContent
            });

            Console.WriteLine("Gui thong tin toi Eoffice nhu sau: " + payload);

            HttpResponseMessage response = await client.PostAsync(config.eoffice.apiUrl, new StringContent(payload, Encoding.UTF8, "application/json"));

            if (response.IsSuccessStatusCode)
            {
                Uri ncrUrl = response.Headers.Location;
                Console.WriteLine("Da gui thong tin toi Eoffice thanh cong");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("Loi!: " + ex.Message);
    }
    #endregion
}
else
{
    Console.WriteLine("Loi: Khong tim thay thu muc da duoc chi dinh");
    Console.ReadLine();
}
Console.WriteLine("Nhan Enter de thoat!");
Console.ReadLine();
driver.Quit();
