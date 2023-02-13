using ASquare.WindowsTaskScheduler;
using ASquare.WindowsTaskScheduler.Models;
using GemBox.Spreadsheet;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TuDongThongBao
{
    public partial class Form1 : Form
    {
        Root config;
        public Form1()
        {
            InitializeComponent();
        }

        private async void Form1_LoadAsync(object sender, EventArgs e)
        {
            try
            {
                config = JsonConvert.DeserializeObject<Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/config.json"));
                Tu.Value = int.Parse(config.alert.fromDay);
                Den.Value = int.Parse(config.alert.toDay);
                DSNhan.Text = string.Join(",",config.alert.users);
                TieuDe.Text = config.alert.message.header;
                NoiDung.Text = config.alert.message.content;
                CuoiTieuDe.Text = config.alert.message.footer;
                EOfficeToken.Text = config.eoffice.token;
                DuongDan.Text = config.download.folderDir;
                ChonDonVi.Checked = config.unitName.isUse == "yes" ? true : false;
                DonVi.Text = config.unitName.choose;
                ChonLoaiThietBi.Checked = config.deviceType.isUse == "yes" ? true : false;
                LoaiThietBi.Text = config.deviceType.choose;
                ChonTenThietBi.Checked = config.deviceName.isUse == "yes" ? true : false;
                TenThietBi.Text = config.deviceName.choose;
                ChonCapDA.Checked = config.levelA.isUse == "yes" ? true : false;
                CapDA.Text = config.levelA.choose;
                ChonNhom.Checked = config.team.isUse == "yes" ? true : false;
                Nhom.Text = config.team.choose;
                ChonDVTT.Checked = config.unit.includeAll == "yes" ? true : false;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
            }
        }

        private void LuuCauHinh_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn lưu cấu hình", "Hỏi", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                config.alert.fromDay = Tu.Value.ToString();
                config.alert.toDay = Den.Value.ToString();
                config.alert.users = DSNhan.Text.Split(',').ToList();
                config.alert.message.header = TieuDe.Text;
                config.alert.message.content = NoiDung.Text;
                config.alert.message.footer = CuoiTieuDe.Text;
                config.eoffice.token = EOfficeToken.Text;
                config.download.folderDir = DuongDan.Text;
                config.unitName.isUse = ChonDonVi.Checked ? "yes" : "no";
                config.unitName.choose = DonVi.Text;
                config.deviceType.isUse = ChonLoaiThietBi.Checked? "yes" : "no";
                config.deviceType.choose = LoaiThietBi.Text;
                config.deviceName.isUse = ChonTenThietBi.Checked ? "yes" : "no";
                config.deviceName.choose = TenThietBi.Text;
                config.levelA.isUse = ChonCapDA.Checked ? "yes" : "no";
                config.levelA.choose = CapDA.Text;
                config.team.isUse = ChonNhom.Checked? "yes" : "no";
                config.team.choose = Nhom.Text;
                config.unit.includeAll = ChonDVTT.Checked ? "yes" : "no";
                File.WriteAllText(Directory.GetCurrentDirectory() + "/config.json", JsonConvert.SerializeObject(config, Formatting.Indented));
                MessageBox.Show("Đã lưu cấu hình");
            }
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void ChonDuongDan_Click(object sender, EventArgs e)
        {
            var folderDlg = new FolderBrowserDialog();
            if (folderDlg.ShowDialog() == DialogResult.OK)
                DuongDan.Text = folderDlg.SelectedPath;
        }

        private async void Chay_Click(object sender, EventArgs e)
        {
            try
            {
                NhatKy.Text = "";
                NhatKy.Text += "Bắt đầu thực hiện";
                var driver = new ChromeDriver();
                NhatKy.Text += "\r\nBắt đầu đọc tệp cấu hình";
                var config = JsonConvert.DeserializeObject<Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/config.json"));
                var dataConfig = JsonConvert.DeserializeObject<DataConfig.Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/dataConfig.json"));
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                if (config == null || dataConfig == null)
                {
                    NhatKy.Text += "\r\nLỗi: Không đọc được tệp cấu hình";
                    return;
                }
                NhatKy.Text += "\r\nĐã đọc xong tệp cấu hình";
                NhatKy.Text += "\r\nĐang lấy dữ liệu một cách tự động hóa";
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
                NhatKy.Text += "\r\nĐang lấy dữ liệu thiết bị";
                Task.Delay(2000).Wait();
                var exportXLSX = driver.FindElement(By.Id(config.exportXLSX.id));
                exportXLSX.Click();
                NhatKy.Text += "\r\nĐang chờ tải tệp dữ liệu thiết bị";
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
                        NhatKy.Text += "\r\nLỗi! Không tìm thấy tệp đã tải về";
                        return;
                    }
                    var lastestFile = (FileInfo)files[0];
                    if (DateTime.Now.Subtract(lastestFile.CreationTime) >= TimeSpan.FromDays(1))
                    {
                        NhatKy.Text += "\r\nLỗi! Tệp đã tải về không phải là tệp mới nhất trong hôm nay";
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
                    NhatKy.Text += "\r\nTong so thiet bi duoc thu thap la: " + deviceList.Count.ToString();
                    NhatKy.Text += "\r\nBat dau loc danh sach cac thiet bi thoa man dien kien";
                    var deviceListNeedInform = deviceList.Where(item => item.RemainingDay != null).Where(item => int.Parse(item.RemainingDay) >= int.Parse(config.alert.fromDay) && int.Parse(item.RemainingDay) <= int.Parse(config.alert.toDay)).ToList();
                    NhatKy.Text += String.Format("\r\nTong so thiet bi thoa man dieu kien tu ngay {0} toi ngay {1}: {2}", config.alert.fromDay, config.alert.toDay, deviceListNeedInform.Count);
                    NhatKy.Text += "\r\nDanh sach gom co: ";
                    foreach (var item in deviceListNeedInform)
                    {
                        NhatKy.Text += String.Format("\r\n{0}\t{1}\t{2}", item.DeviceName, item.LastestUsedDate, item.DeviceStatus);
                    }
                    NhatKy.Text += "\r\nGui thong bao toi Eoffice";
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

                            NhatKy.Text += "\r\nGui thong tin toi Eoffice nhu sau: " + payload;

                            HttpResponseMessage response = await client.PostAsync(config.eoffice.apiUrl, new StringContent(payload, Encoding.UTF8, "application/json"));

                            if (response.IsSuccessStatusCode)
                            {
                                Uri ncrUrl = response.Headers.Location;
                                NhatKy.Text += "Da gui thong tin toi Eoffice thanh cong";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        NhatKy.Text += "\r\nLoi!: " + ex.Message;
                    }
                    #endregion
                }
                else
                {
                    NhatKy.Text += "\r\nLoi: Khong tim thay thu muc da duoc chi dinh";
                }
                driver.Quit();
            }


            catch (Exception ex)
            {
                NhatKy.Text += "\r\n Lỗi: " + ex.Message;
            }
        }

        private void ChonDonVi_CheckedChanged(object sender, EventArgs e)
        {
            DonVi.ReadOnly = !ChonDonVi.Checked;
        }

        private void ChonLoaiThietBi_CheckedChanged(object sender, EventArgs e)
        {
            LoaiThietBi.ReadOnly = !ChonLoaiThietBi.Checked;
        }

        private void ChonTenThietBi_CheckedChanged(object sender, EventArgs e)
        {
            TenThietBi.ReadOnly = !ChonTenThietBi.Checked;
        }

        private void ChonCapDA_CheckedChanged(object sender, EventArgs e)
        {
            CapDA.ReadOnly = !ChonCapDA.Checked;
        }

        private void ChonNhom_CheckedChanged(object sender, EventArgs e)
        {
            Nhom.ReadOnly = !ChonNhom.Checked;
        }

        private async void ChayAn_Click(object sender, EventArgs e)
        {
            try
            {
                NhatKy.Text = "";
                NhatKy.Text += "Bắt đầu thực hiện";
                var service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                var options = new ChromeOptions();
                options.AddArgument("headless");
                var driver = new ChromeDriver(service, options);
                NhatKy.Text += "\r\nBắt đầu đọc tệp cấu hình";
                var config = JsonConvert.DeserializeObject<Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/config.json"));
                var dataConfig = JsonConvert.DeserializeObject<DataConfig.Root>(File.ReadAllText(Directory.GetCurrentDirectory() + "/dataConfig.json"));
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                if (config == null || dataConfig == null)
                {
                    NhatKy.Text += "\r\nLỗi: Không đọc được tệp cấu hình";
                    return;
                }
                NhatKy.Text += "\r\nĐã đọc xong tệp cấu hình";
                NhatKy.Text += "\r\nĐang lấy dữ liệu một cách tự động hóa";
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
                NhatKy.Text += "\r\nĐang lấy dữ liệu thiết bị";
                Task.Delay(2000).Wait();
                var exportXLSX = driver.FindElement(By.Id(config.exportXLSX.id));
                exportXLSX.Click();
                NhatKy.Text += "\r\nĐang chờ tải tệp dữ liệu thiết bị";
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
                        NhatKy.Text += "\r\nLỗi! Không tìm thấy tệp đã tải về";
                        return;
                    }
                    var lastestFile = (FileInfo)files[0];
                    if (DateTime.Now.Subtract(lastestFile.CreationTime) >= TimeSpan.FromDays(1))
                    {
                        NhatKy.Text += "\r\nLỗi! Tệp đã tải về không phải là tệp mới nhất trong hôm nay";
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
                    NhatKy.Text += "\r\nTong so thiet bi duoc thu thap la: " + deviceList.Count.ToString();
                    NhatKy.Text += "\r\nBat dau loc danh sach cac thiet bi thoa man dien kien";
                    var deviceListNeedInform = deviceList.Where(item => item.RemainingDay != null).Where(item => int.Parse(item.RemainingDay) >= int.Parse(config.alert.fromDay) && int.Parse(item.RemainingDay) <= int.Parse(config.alert.toDay)).ToList();
                    NhatKy.Text += String.Format("\r\nTong so thiet bi thoa man dieu kien tu ngay {0} toi ngay {1}: {2}", config.alert.fromDay, config.alert.toDay, deviceListNeedInform.Count);
                    NhatKy.Text += "\r\nDanh sach gom co: ";
                    foreach (var item in deviceListNeedInform)
                    {
                        NhatKy.Text += String.Format("\r\n{0}\t{1}\t{2}", item.DeviceName, item.LastestUsedDate, item.DeviceStatus);
                    }
                    NhatKy.Text += "\r\nGui thong bao toi Eoffice";
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

                            NhatKy.Text += "\r\nGui thong tin toi Eoffice nhu sau: " + payload;

                            HttpResponseMessage response = await client.PostAsync(config.eoffice.apiUrl, new StringContent(payload, Encoding.UTF8, "application/json"));

                            if (response.IsSuccessStatusCode)
                            {
                                Uri ncrUrl = response.Headers.Location;
                                NhatKy.Text += "Da gui thong tin toi Eoffice thanh cong";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        NhatKy.Text += "\r\nLoi!: " + ex.Message;
                    }
                    #endregion
                }
                else
                {
                    NhatKy.Text += "\r\nLoi: Khong tim thay thu muc da duoc chi dinh";
                }
                driver.Quit();
            }


            catch (Exception ex)
            {
                NhatKy.Text += "\r\n Lỗi: " + ex.Message;
            }
        }
    }
}
