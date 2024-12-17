using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace AutoGifteeBox
{
    public partial class fMain : Form
    {
        private bool check;
        private bool isRunning = false;
        private bool stopRequested = false;
        private int lastProcessedRow = 0;

        public fMain()
        {
            InitializeComponent();
        }
        #region License
        private string RunCMD(string cmd)
        {
            Process cmdProcess;
            cmdProcess = new Process();
            cmdProcess.StartInfo.FileName = "cmd.exe";
            cmdProcess.StartInfo.Arguments = "/c " + cmd;
            cmdProcess.StartInfo.RedirectStandardOutput = true;
            cmdProcess.StartInfo.UseShellExecute = false;
            cmdProcess.StartInfo.CreateNoWindow = true;
            cmdProcess.StartInfo.Verb = "runas";
            cmdProcess.Start();
            string output = cmdProcess.StandardOutput.ReadToEnd();
            cmdProcess.WaitForExit();
            if (String.IsNullOrEmpty(output))
                return "";
            return output;
        }

        void CheckLicense()
        {
            string output = RunCMD("wmic diskdrive get serialNumber"); // check số serial ổ cứng
            using (StreamWriter HDD = new StreamWriter("HDD.txt", true))
            {
                HDD.WriteLine(output);
                HDD.Close();
            }
            string[] lines = File.ReadAllLines("HDD.txt");
            File.Delete("HDD.txt");
            string str = Regex.Replace(lines[2], @"\s", ""); // lấy serial đầu tiên

            string outputs = RunCMD("wmic bios get serialnumber"); // check số serial bios
            using (StreamWriter BIOS = new StreamWriter("bios.txt", true))
            {
                BIOS.WriteLine(outputs);
                BIOS.Close();
            }
            string[] liness = File.ReadAllLines("bios.txt");
            File.Delete("bios.txt");
            string strs = Regex.Replace(liness[2], @"\s", ""); // lấy serial đầu tiên

            string keys = string.Concat(strs, str);

            HttpClient httpClient = new HttpClient();
            string text2 = keys;
            string requestUri2 = "https://docs.google.com/spreadsheets/d/1JwQmNaha0kaZzEOCnxjqb3j04zEkZwR6MxAoUyg14Ig/edit?usp=sharing";
            string text3 = httpClient.GetAsync(requestUri2).Result.Content.ReadAsStringAsync().Result.ToString();
            Match match2 = Regex.Match(text3.ToString(), text2 + ".*?(?=ok)");
            bool flag2 = match2 != Match.Empty;
            if (flag2)
            {
                string[] array = match2.ToString().Split(new char[]
                {
                            '|'
                });
                //string siteurlold = "https://mmosorfware.com/time.php";
                //ServicePointManager.Expect100Continue = true;
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                //string htmlold = new System.Net.WebClient().DownloadString(siteurlold);
                //string[] arrayn = htmlold.ToString().Split(new char[]
                //{
                //             '/'
                //});
                //int dayn = Int32.Parse(arrayn[0]);
                //int monthn = Int32.Parse(arrayn[1]);
                //int yearn = Int32.Parse(arrayn[2]);

                DateTime time = DateTime.Now;
                int dayn = time.Day;
                int monthn = time.Month;
                int yearn = time.Year;

                string[] arrays = array[1].ToString().Split(new char[]
               {
                            '/'
               });

                int dayt = Int32.Parse(arrays[0]);
                int montht = Int32.Parse(arrays[1]);
                int yeart = Int32.Parse(arrays[2]);

                System.DateTime now = new System.DateTime(yearn, monthn, dayn);
                System.DateTime then = new System.DateTime(yeart, montht, dayt);
                System.TimeSpan diff1 = then.Subtract(now);


                int days = (int)Math.Ceiling(diff1.TotalDays);

                bool flag3 = days <= 0;
                if (flag3)
                {
                    //MessageBox.Show("Vui lòng liên hệ admin để gia hạn.", "Phần mềm hết hạn" + days, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    //this.tslExp.Text = array[1].ToString();
                    //this.tslRemainTime.Text = "Còn lại: " + days.ToString() + " ngày";
                    Application.Exit();
                }
                else
                {
                    //MessageBox.Show("Đăng Nhập Thành Công !", "Còn lại: " + days + " ngày!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    //this.tslExp.Text = array[1].ToString();
                    //this.tslRemainTime.Text = "Còn lại: " + days.ToString() + " ngày";
                }
            }

            else
            {
                // MessageBox.Show(string.Format("Bạn chưa mua bản quyền tool, vui lòng bấm Ctrl + C và gửi mã \"{0}\" cho chúng tôi để kích hoạt tool, bấm OK để sao chép key!", keys), "Thông báo active bản quyền!", MessageBoxButtons.OK);
                Clipboard.SetText(keys);
                //Environment.Exit(Environment.ExitCode);
                Application.Exit();
            }
        }

        #endregion
        void GetDataFromSheet(string Sheet_ID)
        {
            string sheetUrl = "https://docs.google.com/spreadsheets/d/" + Sheet_ID + "/gviz/tq?tqx=out:csv";

            try
            {
                WebClient client = new WebClient();
                string csvData = client.DownloadString(sheetUrl);

                StringReader reader = new StringReader(csvData);
                DataTable table = new DataTable();
                table.Columns.Add("URL");

                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    int percentIndex = line.IndexOf('%');
                    if (percentIndex != -1)
                    {
                        line = line.Substring(0, percentIndex);
                    }

                    table.Rows.Add(line);
                }

                dgvData.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}");
            }
        }


        void OpenLink(int rowIndex, string link)
        {
            ChromeDriverService cService = ChromeDriverService.CreateDefaultService();
            cService.HideCommandPromptWindow = true;
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--window-size=600,800");
            ChromeDriver driver = new ChromeDriver(cService, options);
            try
            {
                driver.Navigate().GoToUrl(link);

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                try
                {
                    var checkbox = wait.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("MuiIconButton-label")));
                    checkbox.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn1 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div/div[2]/div/button")));
                    btn1.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn2 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div[2]/div/div[2]/div[3]/div[3]")));
                    btn2.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn3 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div/div/div/div/button/img")));
                    btn3.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn4 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div[2]/div/div[2]/button[1]")));
                    Actions actions = new Actions(driver);
                    actions.MoveToElement(btn4);
                    actions.Perform();
                    btn4.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn5 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div[2]/div/div[2]/div/div/div/div[3]/button")));
                    btn5.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                try
                {
                    var btn6 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div[2]/div/div[4]/div/button[1]")));
                    btn6.Click();
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }

                Thread.Sleep(1000);
                try
                {
                    var btn7 = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='root']/div[2]/div/div[2]/div/div/div/div/button")));
                    try
                    {
                        btn7.Click();
                    }
                    catch
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("document.getElementsByClassName('MuiButtonBase-root MuiCardActionArea-root jss46')[0].click();");
                    }
                }
                catch (Exception ex)
                {
                    check = false;
                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                }


                string final_URL = driver.Url;

                if (final_URL.Contains("https://giftee-paypay.e-gift.co/"))
                {
                    check = true;
                }

                Invoke(new Action(() =>
                {
                    dgvData.Rows[rowIndex].Cells["FinalURL"].Value = final_URL;
                }));

                Thread.Sleep(1000);
                driver.Close();
                driver.Quit();
            }
            catch (Exception ex)
            {
                check = false;
                UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
            }
            finally
            {
                driver.Quit();
            }

        }


        private void btnUpdate_Click(object sender, EventArgs e)
        {
            ChromeDriverUpdater.UpdateChromeDriver();
        }

        private void btnGetData_Click(object sender, EventArgs e)
        {
            GetDataFromSheet(txtIDSheet.Text);
            btnRun.Enabled = true;
        }

        private void UpdateStatus(int rowIndex, string status)
        {
            if (dgvData.InvokeRequired)
            {
                dgvData.Invoke(new Action(() =>
                {
                    dgvData.Rows[rowIndex].Cells["Status"].Value = status;
                }));
            }
            else
            {
                dgvData.Rows[rowIndex].Cells["Status"].Value = status;
            }
        }

        private void OpenLinkWithStatus(int rowIndex, string link)
        {
            try
            {
                UpdateStatus(rowIndex, "Đang chạy...");
                OpenLink(rowIndex, link);

                if (check == true)
                {
                    UpdateStatus(rowIndex, "Hoàn thành");
                }

            }
            catch (Exception ex)
            {


            }
        }
        private void RunTaskBasedOnThreads()
        {
            int numberOfThreads = (int)numThreads.Value;
            int rowCount = dgvData.Rows.Count;

            if (numberOfThreads > rowCount)
            {
                MessageBox.Show("Số luồng không thể lớn hơn số dòng dữ liệu!");
                return;
            }

            List<Task> tasks = new List<Task>();
            isRunning = true;

            while (lastProcessedRow < dgvData.Rows.Count)
            {
                if (stopRequested)
                {
                    isRunning = false;
                    break;
                }

                int rowsToProcess = Math.Min(numberOfThreads, dgvData.Rows.Count - lastProcessedRow);

                for (int i = 0; i < rowsToProcess; i++)
                {
                    try
                    {
                        int rowIndex = lastProcessedRow + i;
                        string url = dgvData.Rows[rowIndex].Cells["URL"].Value.ToString().Replace("\"", "");

                        var task = Task.Run(() =>
                        {
                            try
                            {
                                OpenLinkWithStatus(rowIndex, url);
                            }
                            catch (Exception ex)
                            {
                                Invoke(new Action(() =>
                                {
                                    UpdateStatus(rowIndex, $"Lỗi: {ex.Message}");
                                }));
                            }
                        });

                        tasks.Add(task);
                    }
                    catch (Exception ex)
                    {

                    }
                }

                Task.WhenAll(tasks).Wait();
                tasks.Clear();

                Invoke(new Action(() =>
                {
                    for (int i = 0; i < rowsToProcess; i++)
                    {
                        int rowIndex = lastProcessedRow + i;
                        dgvData.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                }));

                lastProcessedRow += rowsToProcess;
            }

            isRunning = false;
        }




        private void btnRun_Click(object sender, EventArgs e)
        {
            if (!isRunning)
            {
                stopRequested = false;
                Task.Run(() => RunTaskBasedOnThreads());
            }
            else
            {

            }
        }

        private void fMain_Load(object sender, EventArgs e)
        {
            CheckLicense();
            btnRun.Enabled = false;
            if (!dgvData.Columns.Contains("Status"))
            {
                DataGridViewTextBoxColumn statusColumn = new DataGridViewTextBoxColumn
                {
                    Name = "Status",
                    HeaderText = "Trạng thái",
                    ReadOnly = true
                };
                dgvData.Columns.Add(statusColumn);
            }

            if (!dgvData.Columns.Contains("FinalURL"))
            {
                DataGridViewTextBoxColumn finalURLColumn = new DataGridViewTextBoxColumn
                {
                    Name = "FinalURL",
                    HeaderText = "Final URL",
                    ReadOnly = true
                };
                dgvData.Columns.Add(finalURLColumn);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Lưu file Excel";
                    saveFileDialog.FileName = "Final_URLs.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        using (ExcelPackage excel = new ExcelPackage())
                        {
                            var worksheet = excel.Workbook.Worksheets.Add("Finished");
                            worksheet.Cells[1, 1].Value = "STT";
                            worksheet.Cells[1, 2].Value = "URL";
                            worksheet.Cells[1, 3].Value = "URL Final";

                            int row = 2;
                            for (int i = 0; i < dgvData.Rows.Count; i++)
                            {
                                if (dgvData.Rows[i].Cells["Status"].Value?.ToString() == "Hoàn thành")
                                {
                                    worksheet.Cells[row, 1].Value = row - 1;
                                    worksheet.Cells[row, 2].Value = dgvData.Rows[i].Cells["URL"].Value?.ToString().Replace("\"", "");
                                    worksheet.Cells[row, 3].Value = dgvData.Rows[i].Cells["FinalURL"].Value?.ToString();
                                    row++;
                                }
                            }

                            FileInfo excelFile = new FileInfo(filePath);
                            excel.SaveAs(excelFile);
                        }

                        MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi: {ex.Message}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (isRunning)
            {
                stopRequested = true;
            }
        }
    }
}
