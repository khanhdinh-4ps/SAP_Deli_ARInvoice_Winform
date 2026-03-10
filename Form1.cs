using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Interop;
namespace SAP_Deli_ARInvoice_Winform
{
    public partial class ARInvoice : Form
    {
        private string currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        private BindingSource bsHeader = new BindingSource();
        private string versionUrl = "https://github.com/khanhdinh-4ps/SAP_Deli_ARInvoice_Winform/releases/latest/download/version.txt";
        private string zipUrl = "https://github.com/khanhdinh-4ps/SAP_Deli_ARInvoice_Winform/releases/latest/download/Debug.zip";
        private Dictionary<int, DataTable> _invoiceLinesCache = new Dictionary<int, DataTable>();
        private DataTable dtHeader = new DataTable();
        private async void CheckForUpdate()
        {
            try
            {
                string serverVersion;

                try
                {
                    // Tải version từ GitHub (async)
                    serverVersion = (await DownloadStringFromUrlAsync(versionUrl)).Trim();
                    if (string.IsNullOrWhiteSpace(serverVersion))
                    {
                        throw new Exception("Không tải được thông tin version.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "Không thể lấy thông tin phiên bản từ server:\n" + ex.Message,
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return;
                }

                if (serverVersion != currentVersion)
                {
                    DialogResult dr = MessageBox.Show(
                        $"Phiên bản hiện tại: {currentVersion}\n" +
                        $"Phiên bản mới: {serverVersion}\n\n" +
                        $"Bạn có muốn cập nhật không?",
                        "Cập nhật phần mềm",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (dr == DialogResult.Yes)
                    {
                        using (var progressForm = new FormUpdateProgress())
                        {
                            progressForm.Show();
                            progressForm.UpdateStatus("Đang tải bản cập nhật...");
                            Application.DoEvents();

                            string tempZip = Path.Combine(Path.GetTempPath(), "Update_" + Guid.NewGuid().ToString("N") + ".zip");

                            try
                            {
                                // Tải file ZIP từ GitHub và hiển thị tiến trình (async)
                                await DownloadFileFromUrlAsync(zipUrl, tempZip, progressForm);
                            }
                            catch (Exception ex)
                            {
                                progressForm.Close();
                                MessageBox.Show(
                                    "Tải file cập nhật thất bại:\n" + ex.Message,
                                    "Lỗi",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                                return;
                            }

                            progressForm.UpdateStatus("Đang cài đặt...");
                            Application.DoEvents();

                            // Gọi Updater
                            string updaterPath = Path.Combine(Application.StartupPath, "Updater.exe");
                            if (!File.Exists(updaterPath))
                            {
                                progressForm.Close();
                                MessageBox.Show("Updater.exe không tìm thấy trong thư mục ứng dụng.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            var psi = new ProcessStartInfo
                            {
                                FileName = updaterPath,
                                Arguments = $"\"{tempZip}\" \"{Application.ExecutablePath}\"",
                                UseShellExecute = true
                            };
                            Process.Start(psi);

                            progressForm.UpdateStatus("Đang khởi động lại ứng dụng...");
                            await Task.Delay(1500);

                            progressForm.Close();
                            Environment.Exit(0);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra cập nhật:\n" + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async Task DownloadFileFromUrlAsync(string url, string destinationPath, FormUpdateProgress progressForm = null)
        {
            using (var client = new HttpClient())
            using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
            {
                response.EnsureSuccessStatusCode();

                long? totalBytes = response.Content.Headers.ContentLength;
                using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                using (var fs = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    byte[] buffer = new byte[81920];
                    long downloadedBytes = 0;
                    int bytesRead;

                    while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
                    {
                        await fs.WriteAsync(buffer, 0, bytesRead).ConfigureAwait(false);
                        downloadedBytes += bytesRead;

                        if (totalBytes.HasValue)
                        {
                            int percent = (int)(downloadedBytes * 100 / totalBytes.Value);
                            // progressForm must be updated on UI thread
                            if (progressForm != null && progressForm.IsHandleCreated)
                                progressForm.Invoke((Action)(() => progressForm.UpdateStatus($"Đang tải bản cập nhật... {percent}%", percent)));
                        }
                    }
                }
            }
        }
        private async Task<string> DownloadStringFromUrlAsync(string url)
        {
            string logFile = Path.Combine(Path.GetTempPath(), "UpdaterDownload.log");
            LogDownload(logFile, $"Downloading text from: {url}");

            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);
                client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

                try
                {
                    var response = await client.GetAsync(url).ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    string content = (await response.Content.ReadAsStringAsync().ConfigureAwait(false)).Trim();
                    LogDownload(logFile, $"Downloaded text content: {content}");

                    return content;
                }
                catch (HttpRequestException ex)
                {
                    LogDownload(logFile, $"HTTP Error: {ex.Message}");
                    throw new Exception($"Không thể tải thông tin version:\n{ex.Message}");
                }
            }
        }
        private bool VerifyZipFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    return false;

                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length < 100)
                    return false;

                // Kiểm tra ZIP signature (PK)
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] header = new byte[4];
                    if (fs.Read(header, 0, 4) == 4)
                    {
                        // ZIP file bắt đầu bằng: 50 4B 03 04 (PK..)
                        if (header[0] == 0x50 && header[1] == 0x4B)
                        {
                            // Thử mở để verify thêm
                            try
                            {
                                fs.Seek(0, SeekOrigin.Begin);
                                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read, true))
                                {
                                    return zip.Entries.Count > 0;
                                }
                            }
                            catch
                            {
                                return false;
                            }
                        }
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }
        private void LogDownload(string logFile, string message)
        {
            try
            {
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, line, Encoding.UTF8);
            }
            catch
            { }
        }
        public ARInvoice()
        {
            InitializeComponent();
            // Keep keyboard event handling
            this.KeyPreview = true;
            this.KeyDown += ARInvoice_KeyDown;
            this.KeyUp += ARInvoice_KeyUp;
            this.tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            // Ensure only tabPage4 is shown: remove all tabs then add tabPage4
            try
            {
                if (tabControl1 != null)
                {
                    tabControl1.TabPages.Clear();
                    if (tabPage4 != null)
                    {
                        tabControl1.TabPages.Add(tabPage4);
                        tabControl1.SelectedTab = tabPage4;
                    }
                }
            }
            catch
            {
                // Fail silently to avoid breaking form initialization
            }
            txtPassword.PasswordChar = 'T';
        }

        private void ARInvoice_Load(object sender, EventArgs e)
        {
            if (btnLogout != null && btnLogout.Visible)
            {
                btnLogout.Visible = false;
                if (btnLogout.Parent != null)
                {
                    btnLogout.Parent.Controls.Remove(btnLogout);
                }
            }
            this.Size = new Size(444, 500);
            this.KeyPreview = true;
            this.KeyDown += ARInvoice_KeyDown;
            this.KeyUp += ARInvoice_KeyUp;
            this.tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            // Ensure only tabPage4 is visible at load time
            try
            {
                if (tabControl1 != null)
                {
                    tabControl1.TabPages.Clear();
                    if (tabPage4 != null)
                    {
                        tabControl1.TabPages.Add(tabPage4);
                        tabControl1.SelectedTab = tabPage4;
                    }
                }
            }
            catch
            {
                // ignore any errors here so form can still load
            }
            txtPassword.PasswordChar = 'T';
            // Start update check in background without blocking UI (fire-and-forget)
            try
            {
                CheckForUpdate();
            }
            catch
            {
                // swallow — CheckForUpdate shows its own errors
            }
        }
        private HashSet<string> _selectedCodes = new HashSet<string>();
        private DataTable dtOriginal;
        private DataTable dtOriginal1;
        private DataTable dtOriginal2;
        private TabPage tabPage3Backup;
        private bool isShortcutHandled = false;
        private TabPage tabPage1Backup;
        private TabPage tabPage2Backup;
        private TabPage tabPage5Backup;
        private bool isLoggedIn = false;
        private void btnLogin_Click(object sender, EventArgs e)
        {
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string userName = null, password = null, bkavUser = null, bkavPass = null;

            try { userName = DecryptAES(ConfigurationManager.AppSettings["UserID"], decryptKey); }
            catch { userName = ConfigurationManager.AppSettings["UserID"]; }

            try { password = DecryptAES(ConfigurationManager.AppSettings["UserPassword"], decryptKey); }
            catch { password = ConfigurationManager.AppSettings["UserPassword"]; }

            string userID = txtUserID.Text.Trim();
            string passwordInput = txtPassword.Text.Trim();

            try { bkavUser = DecryptAES(ConfigurationManager.AppSettings["BkavUser"], decryptKey); }
            catch { bkavUser = ConfigurationManager.AppSettings["BkavUser"]; }

            try { bkavPass = DecryptAES(ConfigurationManager.AppSettings["BkavPassword"], decryptKey); }
            catch { bkavPass = ConfigurationManager.AppSettings["BkavPassword"]; }

            if (string.IsNullOrEmpty(userID) || string.IsNullOrEmpty(passwordInput))
            {
                MessageBox.Show("Vui lòng nhập User ID và Password.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (userID != userName || passwordInput != password)
            {
                if (userID != bkavUser || passwordInput != bkavPass)
                {
                    MessageBox.Show("User ID hoặc Password không đúng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            isLoggedIn = true;
            if (userID == userName && passwordInput == password)
            {
                try
                {
                    if (tabControl1 != null)
                    {
                        tabControl1.TabPages.Clear();

                        var page1 = tabPage1 ?? tabPage1Backup;
                        //var page5 = tabPage5 ?? tabPage5Backup;
                        var page2 = tabPage2 ?? tabPage2Backup;

                        if (page1 != null && !tabControl1.TabPages.Contains(page1))
                            tabControl1.TabPages.Add(page1);

/*                        if (page5 != null && !tabControl1.TabPages.Contains(page5))
                            tabControl1.TabPages.Add(page5);*/

                        if (page2 != null && !tabControl1.TabPages.Contains(page2))
                            tabControl1.TabPages.Add(page2);

                        if (tabControl1.TabPages.Count > 0)
                            tabControl1.SelectedIndex = 0;
                    }

                    this.Size = new Size(1119, 666);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi hiển thị tab: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (userID == bkavUser && passwordInput == bkavPass)
            {
                if (tabPage6 != null && !tabControl1.TabPages.Contains(tabPage6))
                {
                    tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tabPage6);
                }
                if (tabPage7 != null && !tabControl1.TabPages.Contains(tabPage7))
                {
                    tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tabPage7);
                }
                tabControl1.TabPages.Remove(tabPage4);
                this.Size = new Size(1119, 666);
            }
            txtUserID.Text = string.Empty;
            txtPassword.Text = string.Empty;
            MessageBox.Show("Đăng nhập thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void ARInvoice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.Shift && e.KeyCode == Keys.N)
            {
                if (!isShortcutHandled)
                {
                    isShortcutHandled = true;
                    try
                    {
                        if (tabControl1 == null) return;

                        var page = tabPage3 ?? tabPage3Backup;

                        if (page == null)
                        {
                            return;
                        }

                        if (tabControl1.TabPages.Contains(page))
                        {
                            tabPage3Backup = page;
                            tabControl1.TabPages.Remove(page);
                        }
                        else
                        {
                            if (!tabControl1.TabPages.Contains(page))
                            {
                                tabControl1.TabPages.Add(page);
                                tabControl1.SelectedTab = page;
                            }
                        }
                    }
                    catch (ArgumentNullException)
                    {
                        
                    }
                    catch
                    {
                        
                    }
                }
            }
            if (e.Control && e.Shift && e.KeyCode == Keys.B)
            {
                if (btnLogout != null)
                {
                    btnLogout.Visible = !btnLogout.Visible;
                    if (btnLogout.Visible)
                    {
                        try
                        {
                            if (btnLogout.Parent == null || btnLogout.Parent != this)
                            {
                                if (btnLogout.Parent != null)
                                {
                                    var prevParent = btnLogout.Parent;
                                    prevParent.Controls.Remove(btnLogout);
                                }
                                this.Controls.Add(btnLogout);
                                btnLogout.BringToFront();
                                btnLogout.Anchor = AnchorStyles.Top | AnchorStyles.Left; btnLogout.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                            }
                        }
                        catch
                        {
                            
                        }
                    }
                }
            }
        }
        private void ARInvoice_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.N)
            {
                isShortcutHandled = false;
            }
        }
        private Company ConnectToSAP()
        {
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string server = null, companyDB = null, userName = null, password = null, dbUserName = null, dbPassword = null, licenseServer = null;

            try { server = DecryptAES(ConfigurationManager.AppSettings["SAP_Server"], decryptKey); }
            catch { server = ConfigurationManager.AppSettings["SAP_Server"]; }

            try { companyDB = DecryptAES(ConfigurationManager.AppSettings["SAP_CompanyDB"], decryptKey); }
            catch { companyDB = ConfigurationManager.AppSettings["SAP_CompanyDB"]; }

            try { userName = DecryptAES(ConfigurationManager.AppSettings["SAP_UserName"], decryptKey); }
            catch { userName = ConfigurationManager.AppSettings["SAP_UserName"]; }

            try { password = DecryptAES(ConfigurationManager.AppSettings["SAP_Password"], decryptKey); }
            catch { password = ConfigurationManager.AppSettings["SAP_Password"]; }

            try { dbUserName = DecryptAES(ConfigurationManager.AppSettings["SAP_DbUserName"], decryptKey); }
            catch { dbUserName = ConfigurationManager.AppSettings["SAP_DbUserName"]; }

            try { dbPassword = DecryptAES(ConfigurationManager.AppSettings["SAP_DbPassword"], decryptKey); }
            catch { dbPassword = ConfigurationManager.AppSettings["SAP_DbPassword"]; }

            try { licenseServer = DecryptAES(ConfigurationManager.AppSettings["SAP_LicenseServer"], decryptKey); }
            catch { licenseServer = ConfigurationManager.AppSettings["SAP_LicenseServer"]; }
            try
            {
                Company company = new Company
                {
                    Server = server,
                    CompanyDB = companyDB,
                    UserName = userName,
                    Password = password,
                    DbServerType = BoDataServerTypes.dst_MSSQL2017,
                    LicenseServer = licenseServer,
                    DbUserName = dbUserName,
                    DbPassword = dbPassword
                };
                int result = company.Connect();
                if (result == 0)
                {
                    MessageBox.Show("Connection Checking Succesful");
                }
                if (result != 0)
                {
                    MessageBox.Show("Kết nối SAP B1 thất bại: " + company.GetLastErrorDescription());
                    return null;
                }
                if (company.Connected)
                    company.Disconnect();
                return company;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi kết nối SAP: {ex.Message}");
                return null;
            }
        }
        private Company ConnectToSAP2()
        {
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string server = null, companyDB = null, userName = null, password = null, dbUserName = null, dbPassword = null, licenseServer = null;

            try { server = DecryptAES(ConfigurationManager.AppSettings["SAP_Server"], decryptKey); }
            catch { server = ConfigurationManager.AppSettings["SAP_Server"]; }

            try { companyDB = DecryptAES(ConfigurationManager.AppSettings["SAP_CompanyDB"], decryptKey); }
            catch { companyDB = ConfigurationManager.AppSettings["SAP_CompanyDB"]; }

            try { userName = DecryptAES(ConfigurationManager.AppSettings["SAP_UserName"], decryptKey); }
            catch { userName = ConfigurationManager.AppSettings["SAP_UserName"]; }

            try { password = DecryptAES(ConfigurationManager.AppSettings["SAP_Password"], decryptKey); }
            catch { password = ConfigurationManager.AppSettings["SAP_Password"]; }

            try { dbUserName = DecryptAES(ConfigurationManager.AppSettings["SAP_DbUserName"], decryptKey); }
            catch { dbUserName = ConfigurationManager.AppSettings["SAP_DbUserName"]; }

            try { dbPassword = DecryptAES(ConfigurationManager.AppSettings["SAP_DbPassword"], decryptKey); }
            catch { dbPassword = ConfigurationManager.AppSettings["SAP_DbPassword"]; }

            try { licenseServer = DecryptAES(ConfigurationManager.AppSettings["SAP_LicenseServer"], decryptKey); }
            catch { licenseServer = ConfigurationManager.AppSettings["SAP_LicenseServer"]; }

            try
            {
                Company company = new Company
                {
                    Server = server,
                    CompanyDB = companyDB,
                    UserName = userName,
                    Password = password,
                    DbServerType = BoDataServerTypes.dst_MSSQL2017,
                    LicenseServer = licenseServer,
                    DbUserName = dbUserName,
                    DbPassword = dbPassword
                };
                int result = company.Connect();
                if (result != 0)
                {
                    MessageBox.Show("Kết nối SAP B1 thất bại: " + company.GetLastErrorDescription());
                    return null;
                }
                return company;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi kết nối SAP: {ex.Message}");
                return null;
            }
        }
        private Company ConnectToSAP3()
        {
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string server = null, companyDB = null, userName = null, password = null, dbUserName = null, dbPassword = null, licenseServer = null;

            try { server = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_Server"], decryptKey); }
            catch { server = ConfigurationManager.AppSettings["KitchenSAP_Server"]; }

            try { companyDB = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_CompanyDB"], decryptKey); }
            catch { companyDB = ConfigurationManager.AppSettings["KitchenSAP_CompanyDB"]; }

            try { userName = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_UserName"], decryptKey); }
            catch { userName = ConfigurationManager.AppSettings["KitchenSAP_UserName"]; }

            try { password = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_Password"], decryptKey); }
            catch { password = ConfigurationManager.AppSettings["KitchenSAP_Password"]; }

            try { dbUserName = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_DbUserName"], decryptKey); }
            catch { dbUserName = ConfigurationManager.AppSettings["KitchenSAP_DbUserName"]; }

            try { dbPassword = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_DbPassword"], decryptKey); }
            catch { dbPassword = ConfigurationManager.AppSettings["KitchenSAP_DbPassword"]; }

            try { licenseServer = DecryptAES(ConfigurationManager.AppSettings["KitchenSAP_LicenseServer"], decryptKey); }
            catch { licenseServer = ConfigurationManager.AppSettings["KitchenSAP_LicenseServer"]; }

            try
            {
                Company company = new Company
                {
                    Server = server,
                    CompanyDB = companyDB,
                    UserName = userName,
                    Password = password,
                    DbServerType = BoDataServerTypes.dst_MSSQL2017,
                    LicenseServer = licenseServer,
                    DbUserName = dbUserName,
                    DbPassword = dbPassword
                };
                int result = company.Connect();
                if (result != 0)
                {
                    MessageBox.Show("Kết nối SAP B1 thất bại: " + company.GetLastErrorDescription());
                    return null;
                }
                return company;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi kết nối SAP: {ex.Message}");
                return null;
            }
        }
        public static SqlConnection ConnectToSQL()
        {
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string server = null, user = null, password = null, database = null;

            var decryptor = new ARInvoice();

            try { server = decryptor.DecryptAES(ConfigurationManager.AppSettings["SAP_Server4p"], decryptKey); }
            catch { server = ConfigurationManager.AppSettings["SAP_Server4p"]; }

            try { user = decryptor.DecryptAES(ConfigurationManager.AppSettings["SAP_UserName4p"], decryptKey); }
            catch { user = ConfigurationManager.AppSettings["SAP_UserName4p"]; }

            try { password = decryptor.DecryptAES(ConfigurationManager.AppSettings["SAP_Password4p"], decryptKey); }
            catch { password = ConfigurationManager.AppSettings["SAP_Password4p"]; }

            try { database = decryptor.DecryptAES(ConfigurationManager.AppSettings["SAP_CompanyDB4p"], decryptKey); }
            catch { database = ConfigurationManager.AppSettings["SAP_CompanyDB4p"]; }

            string connectionString = $"Server={server};Database={database};User Id={user};Password={password};";

            var conn = new System.Data.SqlClient.SqlConnection(connectionString);
            try
            {
                conn.Open();
            }
            catch
            {
                conn = null;
            }
            return conn;
        }
        private string EncryptAES(string plainText, string key)
        {
            using (Aes aes = Aes.Create())
            {
                var keyBytes = new Rfc2898DeriveBytes(key, Encoding.UTF8.GetBytes("s@ltValue"), 1000);
                aes.Key = keyBytes.GetBytes(32);
                aes.IV = keyBytes.GetBytes(16);

                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    byte[] inputBytes = Encoding.UTF8.GetBytes(plainText);
                    cs.Write(inputBytes, 0, inputBytes.Length);
                    cs.Close();
                    return Convert.ToBase64String(ms.ToArray());
                }
            }
        }
        private string DecryptAES(string cipherText, string key)
        {
            using (Aes aes = Aes.Create())
            {
                var keyBytes = new Rfc2898DeriveBytes(key, Encoding.UTF8.GetBytes("s@ltValue"), 1000);
                aes.Key = keyBytes.GetBytes(32);
                aes.IV = keyBytes.GetBytes(16);

                using (var ms = new MemoryStream())
                using (var cs = new CryptoStream(ms, aes.CreateDecryptor(), CryptoStreamMode.Write))
                {
                    byte[] cipherBytes = Convert.FromBase64String(cipherText);
                    cs.Write(cipherBytes, 0, cipherBytes.Length);
                    cs.Close();
                    return Encoding.UTF8.GetString(ms.ToArray());
                }
            }
        }
        private void btnEncrypt_Click(object sender, EventArgs e)
        {
            string key = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string input = txtInput.Text;
            try
            {
                string encrypted = EncryptAES(input, key);
                txtOutput.Text = encrypted;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Encryption failed: " + ex.Message);
            }
        }
        private void btnDecrypt_Click(object sender, EventArgs e)
        {
            string key = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string cipherText = txtOutput.Text;
            try
            {
                string decrypted = DecryptAES(cipherText, key);
                txtInput.Text = decrypted;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Decryption failed: " + ex.Message);
            }
        }
        private void btnCheck_connection_Click(object sender, EventArgs e)
        {
            ConnectToSAP();
        }
        private bool IsDeliveryOpen(Company company, int docEntry)
        {
            Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                // Truy vấn trực tiếp để lấy trạng thái mới nhất từ Database
                string sql = $"SELECT DocStatus FROM ODLN WHERE DocEntry = {docEntry}";
                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    string status = rs.Fields.Item("DocStatus").Value.ToString();
                    return status == "O"; // 'O' là Open, 'C' là Closed
                }
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); // Giải phóng RAM
            }
        }
        private async void btnCopyToInvoice_Click(object sender, EventArgs e)
        {
            // BƯỚC 1: LẤY DỮ LIỆU TỪ UI (Phải làm ở UI Thread)
            var selectedRowsData = dtGridview_ODLN.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells["Select"].Value != null && (bool)row.Cells["Select"].Value == true)
                .Select(row => new {
                    CardCode = row.Cells["CardCode"].Value.ToString(),
                    DocEntry = Convert.ToInt32(row.Cells["DocEntry"].Value),
                    SlpCode = Convert.ToInt32(row.Cells["SlpCode"]?.Value ?? -1),
                    Active = row.Cells["Active"]?.Value?.ToString()
                })
                .ToList();

            if (selectedRowsData.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một phiếu Delivery.");
                return;
            }

            // BƯỚC 2: CHUẨN BỊ UI
            this.Cursor = Cursors.WaitCursor;
            btnCopyToInvoice.Enabled = false;

            Company company = ConnectToSAP2(); // Kết nối SAP
            if (company == null)
            {
                this.Cursor = Cursors.Default;
                btnCopyToInvoice.Enabled = true;
                return;
            }

            // BƯỚC 3: XỬ LÝ NẶNG Ở BACKGROUND
            var resultSummary = await Task.Run(() =>
            {
                var successList = new List<string>();
                var errorList = new List<string>();
                var successDocEntries = new List<int>();

                foreach (var item in selectedRowsData)
                {
                    if (!IsDeliveryOpen(company, item.DocEntry))
                    {
                        errorList.Add($"Delivery {item.DocEntry}: Phiếu đã bị đóng hoặc hủy trong SAP.");
                        continue; // Bỏ qua phiếu này và sang phiếu tiếp theo
                    }
                    string cardCode = item.CardCode;
                    // Mỗi lần chạy chỉ truyền vào 1 DocEntry duy nhất trong List
                    List<int> docEntries = new List<int> { item.DocEntry };

                    // Xác định nhân viên kinh doanh cho từng dòng
                    int finalSlp = (item.Active == "Y" || item.Active == "TRUE") ? item.SlpCode : -1;

                    // Gọi hàm tạo Invoice (lúc này List docEntries chỉ có 1 phần tử)
                    var res = CreateARInvoiceInternal(company, cardCode, docEntries, finalSlp);

                    if (res.Success)
                    {
                        // Ghi nhận số phiếu Delivery cũ và số Invoice mới
                        successList.Add($"Delivery {item.DocEntry} -> Invoice {res.NewDocEntry}");
                        successDocEntries.Add(item.DocEntry);
                    }
                    else
                    {
                        errorList.Add($"Delivery {item.DocEntry} (Khách: {cardCode}): {res.ErrorMessage}");
                    }
                }
                return new { successList, errorList, successDocEntries };
            });

            // BƯỚC 4: KẾT THÚC VÀ HIỂN THỊ KẾT QUẢ
            if (company.Connected) company.Disconnect();
            this.Cursor = Cursors.Default;
            btnCopyToInvoice.Enabled = true;
            RefreshGridViewAfterSuccess(resultSummary.successDocEntries);
            ShowFinalReport(resultSummary.successList, resultSummary.errorList);
        }
        private void ShowFinalReport(List<string> successList, List<string> errorList)
        {
            StringBuilder sb = new StringBuilder();
            int total = successList.Count + errorList.Count;

            sb.AppendLine($"--- TỔNG KẾT XỬ LÝ ({DateTime.Now:HH:mm:ss}) ---");
            sb.AppendLine($"Tổng số nhóm khách hàng: {total}");
            sb.AppendLine($"Thành công: {successList.Count}");
            sb.AppendLine($"Thất bại: {errorList.Count}");
            sb.AppendLine(new string('-', 40));

            if (successList.Count > 0)
            {
                sb.AppendLine("✅ DANH SÁCH THÀNH CÔNG (A/R Invoice):");
                foreach (var item in successList)
                {
                    sb.AppendLine($"  • {item}");
                }
            }

            if (errorList.Count > 0)
            {
                if (successList.Count > 0) sb.AppendLine(); // Tạo khoảng trống giữa 2 phần
                sb.AppendLine("❌ DANH SÁCH LỖI:");
                foreach (var err in errorList)
                {
                    sb.AppendLine($"  • {err}");
                }
            }

            // Hiển thị Icon dựa trên kết quả
            MessageBoxIcon icon = errorList.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning;
            if (successList.Count == 0 && errorList.Count > 0) icon = MessageBoxIcon.Error;

            MessageBox.Show(sb.ToString(), "Kết quả Copy To Invoice", MessageBoxButtons.OK, icon);
        }
        private void RefreshGridViewAfterSuccess(List<int> successfulDocEntries)
        {
            if (successfulDocEntries == null || successfulDocEntries.Count == 0) return;

            // Find the underlying DataTable whether the grid is bound directly or via a BindingSource
            DataTable sourceDt = null;

            if (dtGridview_ODLN?.DataSource is DataTable directDt)
            {
                sourceDt = directDt;
            }
            else if (dtGridview_ODLN?.DataSource is BindingSource bs && bs.DataSource is DataTable bsDt)
            {
                sourceDt = bsDt;
            }
            else if (bsHeader?.DataSource is DataTable hdrDt)
            {
                sourceDt = hdrDt;
            }
            else if (this.dtHeader != null)
            {
                sourceDt = this.dtHeader;
            }

            if (sourceDt != null)
            {
                for (int i = sourceDt.Rows.Count - 1; i >= 0; i--)
                {
                    if (int.TryParse(sourceDt.Rows[i]["DocEntry"]?.ToString(), out int entry) &&
                        successfulDocEntries.Contains(entry))
                    {
                        sourceDt.Rows.RemoveAt(i);
                    }
                }
                sourceDt.AcceptChanges();
            }

            // 2. Clear cache entries
            foreach (int entry in successfulDocEntries)
            {
                if (_invoiceLinesCache.ContainsKey(entry))
                {
                    _invoiceLinesCache.Remove(entry);
                }
            }

            // 3. Clear detail grid if selected row was removed
            dtGridview_DLN1.DataSource = null;

            // UI feedback
            lblStatus.Text = $"Đã cập nhật giao diện: Đã xử lý {successfulDocEntries.Count} phiếu.";
        }
        private void KitRefreshGridViewAfterSuccess(List<int> successfulDocEntries)
        {
            if (successfulDocEntries == null || successfulDocEntries.Count == 0) return;

            DataTable sourceDt = null;

            if (KitdtGridview_ODLN?.DataSource is DataTable directDt)
            {
                sourceDt = directDt;
            }
            else if (KitdtGridview_ODLN?.DataSource is BindingSource bs && bs.DataSource is DataTable bsDt)
            {
                sourceDt = bsDt;
            }
            else if (bsHeader?.DataSource is DataTable hdrDt)
            {
                sourceDt = hdrDt;
            }
            else if (this.dtHeader != null)
            {
                sourceDt = this.dtHeader;
            }

            if (sourceDt != null)
            {
                for (int i = sourceDt.Rows.Count - 1; i >= 0; i--)
                {
                    if (int.TryParse(sourceDt.Rows[i]["DocEntry"]?.ToString(), out int entry) &&
                        successfulDocEntries.Contains(entry))
                    {
                        sourceDt.Rows.RemoveAt(i);
                    }
                }
                sourceDt.AcceptChanges();
            }

            foreach (int entry in successfulDocEntries)
            {
                if (_invoiceLinesCache.ContainsKey(entry))
                {
                    _invoiceLinesCache.Remove(entry);
                }
            }

            KitdtGridview_DLN1.DataSource = null;

            KitlblStatus.Text = $"Đã cập nhật giao diện: Đã xử lý {successfulDocEntries.Count} phiếu.";
        }
        // Hàm bổ trợ tạo Invoice - Tối ưu bằng cách giải phóng tài nguyên COM ngay lập tức
        private (bool Success, string NewDocEntry, string ErrorMessage) CreateARInvoiceInternal(Company company, string cardCode, List<int> docEntries, int salesPersonCode)
        {
            Documents vInvoice = null;
            try
            {
                vInvoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
                vInvoice.CardCode = cardCode;
                vInvoice.DocDate = DateTime.Now;
                if (salesPersonCode > 0) vInvoice.SalesPersonCode = salesPersonCode;

                int lineCount = 0;
                foreach (int docEntry in docEntries)
                {
                    // Sử dụng TryGetValue để truy cập Cache nhanh hơn
                    if (!_invoiceLinesCache.TryGetValue(docEntry, out DataTable dtLines)) continue;

                    foreach (DataRow row in dtLines.Rows)
                    {
                        if (row["LineStatus"]?.ToString() == "C") continue;

                        if (lineCount > 0) vInvoice.Lines.Add();
                        vInvoice.Lines.BaseType = 15; // oDeliveryNotes
                        vInvoice.Lines.BaseEntry = docEntry;
                        vInvoice.Lines.BaseLine = Convert.ToInt32(row["LineNum"]);

                        decimal netQty = Convert.ToDecimal(row["NetQuantity"] ?? 0);
                        if (netQty > 0) vInvoice.Lines.Quantity = (double)netQty;

                        lineCount++;
                    }
                }

                if (lineCount == 0) return (false, null, "Không có dòng mặt hàng.");

                if (vInvoice.Add() != 0) return (false, null, company.GetLastErrorDescription());

                return (true, company.GetNewObjectKey(), null);
            }
            finally
            {
                // QUAN TRỌNG: Giải phóng bộ nhớ COM để tránh lag server
                if (vInvoice != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(vInvoice);
            }
        }
        private int DetermineInvoiceSalesPerson(IEnumerable<DataGridViewRow> rows)
        {
            try
            {
                var slpSet = new HashSet<int>();
                // Map SlpCode -> active string (take the first non-null Active for that Slp)
                var slpActive = new Dictionary<int, string>();

                foreach (var row in rows)
                {
                    object slpObj = null;
                    object activeObj = null;

                    if (row.Cells["SlpCode"] != null)
                        slpObj = row.Cells["SlpCode"].Value;

                    if (row.Cells["Active"] != null)
                        activeObj = row.Cells["Active"].Value;

                    if (slpObj == null || string.IsNullOrWhiteSpace(slpObj.ToString()))
                        continue;

                    if (int.TryParse(slpObj.ToString(), out int slp))
                    {
                        slpSet.Add(slp);

                        if (!slpActive.ContainsKey(slp))
                        {
                            slpActive[slp] = activeObj?.ToString();
                        }
                    }
                }

                if (slpSet.Count != 1)
                    return 0;

                int singleSlp = slpSet.First();

                string activeStr = null;
                if (slpActive.TryGetValue(singleSlp, out activeStr))
                {
                    activeStr = (activeStr ?? "").Trim();
                }

                if (!string.IsNullOrEmpty(activeStr) &&
                    (string.Equals(activeStr, "Y", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(activeStr, "1", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(activeStr, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(activeStr, "T", StringComparison.OrdinalIgnoreCase)))
                {
                    return singleSlp;
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Kiểm tra cột Checkbox "Chọn"
                if (!dtGridview_ODLN.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn selectColumn = new DataGridViewCheckBoxColumn();
                    selectColumn.Name = "Select";
                    selectColumn.HeaderText = "Chọn";
                    selectColumn.Width = 65;
                    dtGridview_ODLN.Columns.Insert(0, selectColumn);
                }

                // 2. Kiểm tra điều kiện đầu vào
                if (string.IsNullOrWhiteSpace(txtCustomer.Text) && string.IsNullOrWhiteSpace(txtCustomer2.Text))
                {
                    MessageBox.Show("Vui lòng nhập tên khách hàng hoặc mã khách hàng để tìm kiếm.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 3. Xóa Cache cũ trước khi nạp dữ liệu mới
                _invoiceLinesCache.Clear();

                // 4. Kết nối SAP
                Company company = ConnectToSAP2();
                if (company == null) return;

                // 5. Chuẩn bị biến lọc dữ liệu
                string cardName = txtCustomer.Text.Trim().Replace("'", "''");
                string cardCode = txtCustomer2.Text.Trim().Replace("'", "''");
                string fromDate = fromdatePick.Value.ToString("yyyy-MM-dd");
                string toDate = todatePick.Value.ToString("yyyy-MM-dd");

                string statusFilter = "";
                if (checkBox_Open.Checked && !checkBox_Closed.Checked)
                    statusFilter = "AND DocStatus = 'O'";
                else if (!checkBox_Open.Checked && checkBox_Closed.Checked)
                    statusFilter = "AND DocStatus = 'C'";

                string cardCodeFilter = string.IsNullOrWhiteSpace(cardCode) ? "" : $"T0.CardCode LIKE '%{cardCode}%'";
                string cardNameFilter = string.IsNullOrWhiteSpace(cardName) ? "" : $"T0.CardName LIKE N'%{cardName}%'";

                string whereFilter = "";
                if (!string.IsNullOrEmpty(cardCodeFilter) && !string.IsNullOrEmpty(cardNameFilter))
                    whereFilter = $"({cardCodeFilter} AND {cardNameFilter})";
                else if (!string.IsNullOrEmpty(cardCodeFilter))
                    whereFilter = cardCodeFilter;
                else if (!string.IsNullOrEmpty(cardNameFilter))
                    whereFilter = cardNameFilter;

                try
                {
                    string queryHeader = $@"
                    SELECT 
                        ROW_NUMBER() OVER (ORDER BY T0.DocEntry) AS STT,
                        T0.DocEntry, T0.CardCode, T0.CardName, T0.DocStatus, T0.DocType, 
                        CONVERT(varchar(10), T0.DocDate, 103) AS DocDate, 
                        CONVERT(varchar(10), T0.DocDueDate, 103) AS DocDueDate,  
                        T0.DocTotal, T0.SlpCode, S.Active,
                        ISNULL(R0.DocEntry, 0) AS ReturnDocEntry
                    FROM ODLN T0 WITH(NOLOCK)
                    LEFT JOIN OCRD C WITH(NOLOCK) ON C.CardCode = T0.CardCode
                    LEFT JOIN OSLP S WITH(NOLOCK) ON S.SlpCode = T0.SlpCode
                    OUTER APPLY (
                        SELECT TOP 1 R0.DocEntry
                        FROM ORDN R0 WITH(NOLOCK)
                        INNER JOIN RDN1 R1 ON R1.DocEntry = R0.DocEntry
                        WHERE R1.BaseEntry = T0.DocEntry AND R0.CANCELED = 'N'
                        ORDER BY R0.DocEntry
                    ) R0
                    WHERE {whereFilter} 
                      AND T0.DocDate BETWEEN '{fromDate}' AND '{toDate}' 
                      AND ISNULL(C.FrozenFor, 'N') <> 'Y'
                      {statusFilter}
                    ORDER BY T0.DocEntry";

                    Recordset rsHeader = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsHeader.DoQuery(queryHeader);

                    // Nạp dữ liệu vào DataTable cấp Class
                    dtHeader = ConvertRecordsetToDataTable(rsHeader);

                    // Gán DataTable vào BindingSource
                    bsHeader.DataSource = dtHeader;

                    // Gán BindingSource vào DataGridView
                    dtGridview_ODLN.DataSource = bsHeader;

                    // --- BƯỚC 7: QUERY TOÀN BỘ CHI TIẾT (LINES) VÀO CACHE ---
                    // Sử dụng INNER JOIN để lọc các Lines thuộc về các Headers ở trên
                    string queryAllLines = $@"
                    WITH ReturnSummary AS (
                        SELECT R1.BaseEntry, R1.BaseLine, SUM(R1.Quantity) AS ReturnQty
                        FROM RDN1 R1 WITH(NOLOCK)
                        INNER JOIN ORDN R0 WITH(NOLOCK) ON R0.DocEntry = R1.DocEntry
                        WHERE R0.CANCELED = 'N'
                        GROUP BY R1.BaseEntry, R1.BaseLine
                    ),
                    InvoiceSummary AS (
                        SELECT I1.BaseEntry, I1.BaseLine, SUM(I1.Quantity) AS InvoicedQty, MAX(I1.DocEntry) AS InvDocEntry
                        FROM INV1 I1 WITH(NOLOCK)
                        INNER JOIN OINV I0 WITH(NOLOCK) ON I0.DocEntry = I1.DocEntry
                        WHERE I0.CANCELED = 'N' AND I1.BaseType = 15
                        GROUP BY I1.BaseEntry, I1.BaseLine
                    )
                    SELECT 
                        T0.DocEntry, T0.LineNum, T0.ItemCode, T0.Dscription, T0.Quantity, 
                        ISNULL(TR.ReturnQty, 0) AS ReturnQty, 
                        ISNULL(TI.InvoicedQty, 0) AS InvoicedQty,
                        (T0.Quantity - ISNULL(TR.ReturnQty, 0) - ISNULL(TI.InvoicedQty, 0)) AS NetQuantity,
                        T0.Price, T0.LineTotal, T0.VatGroup, T0.GTotal, T0.WhsCode, 
                        T0.AcctCode, T0.LineStatus, T0.TrgetEntry,
                        ISNULL(TI.InvDocEntry, 0) AS InvDocEntry
                    FROM DLN1 T0 WITH(NOLOCK)
                    INNER JOIN ODLN H WITH(NOLOCK) ON T0.DocEntry = H.DocEntry
                    LEFT JOIN ReturnSummary TR ON TR.BaseEntry = T0.DocEntry AND TR.BaseLine = T0.LineNum
                    LEFT JOIN InvoiceSummary TI ON TI.BaseEntry = T0.DocEntry AND TI.BaseLine = T0.LineNum
                    WHERE {whereFilter.Replace("T0.CardCode", "H.CardCode").Replace("T0.CardName", "H.CardName")} 
                      AND H.DocDate BETWEEN '{fromDate}' AND '{toDate}' 
                      {statusFilter.Replace("DocStatus", "H.DocStatus")}";

                    Recordset rsLines = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsLines.DoQuery(queryAllLines);
                    DataTable dtAllLines = ConvertRecordsetToDataTable(rsLines);

                    // --- BƯỚC 8: PHÂN TÁCH DỮ LIỆU VÀO DICTIONARY (CACHE) ---
                    if (dtAllLines.Rows.Count > 0)
                    {
                        // Group dữ liệu theo DocEntry bằng LINQ
                        var grouped = dtAllLines.AsEnumerable().GroupBy(r => Convert.ToInt32(r["DocEntry"]));
                        foreach (var group in grouped)
                        {
                            _invoiceLinesCache[group.Key] = group.CopyToDataTable();
                        }
                    }

                    MessageBox.Show($"Tìm thấy {dtGridview_ODLN.Rows.Count} hóa đơn.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    if (company.Connected) company.Disconnect();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dtGridview_ODLN_SelectionChanged(object sender, EventArgs e)
        {
            if (dtGridview_ODLN.SelectedRows.Count == 0) return;
            var selectedRow = dtGridview_ODLN.SelectedRows[0];
            if (selectedRow.Cells["DocEntry"].Value == null) return;
            int docEntry = Convert.ToInt32(selectedRow.Cells["DocEntry"].Value);
            if (_invoiceLinesCache.ContainsKey(docEntry))
            {
                dtGridview_DLN1.DataSource = _invoiceLinesCache[docEntry];
            }
            else
            {
                // Trường hợp hóa đơn không có dòng nào (hiếm gặp)
                dtGridview_DLN1.DataSource = null;
            }
            FormatDetailGridView();
        }
        private DataTable ConvertRecordsetToDataTable(Recordset recordset)
        {
            DataTable dt = new DataTable();

            // 1. Tạo các cột dựa trên Fields trong Recordset
            for (int i = 0; i < recordset.Fields.Count; i++)
            {
                dt.Columns.Add(recordset.Fields.Item(i).Name);
            }

            // 2. Đổ dữ liệu vào DataTable
            while (!recordset.EoF)
            {
                DataRow row = dt.NewRow();
                for (int i = 0; i < recordset.Fields.Count; i++)
                {
                    row[i] = recordset.Fields.Item(i).Value;
                }
                dt.Rows.Add(row);
                recordset.MoveNext();
            }

            return dt;
        }
        private void FormatDetailGridView()
        {
            if (dtGridview_DLN1.Columns.Count == 0) return;

            Dictionary<string, string> headers = new Dictionary<string, string>
            {
                { "Dscription", "Item Name" },
                { "AcctCode", "G/L Account" },
                { "TrgetEntry", "Return DocEntry" },
                { "InvDocEntry", "Invoice DocEntry" }, // Cột mới thêm
                { "InvoicedQty", "Invoiced Qty" },     // Cột mới thêm
                { "VatGroup", "Tax Code" },
                { "GTotal", "Gross Total" },
                { "NetQuantity", "Remaining Qty" }      // Đổi tên cho rõ nghĩa
            };

            foreach (var header in headers)
            {
                if (dtGridview_DLN1.Columns.Contains(header.Key))
                {
                    dtGridview_DLN1.Columns[header.Key].HeaderText = header.Value;
                }
            }   

            // Cập nhật các cột số lượng cần định dạng
            string[] numericColumns = { "Quantity", "ReturnQty", "InvoicedQty", "NetQuantity", "Price", "LineTotal", "GTotal" };
            foreach (string colName in numericColumns)
            {
                if (dtGridview_DLN1.Columns.Contains(colName))
                {
                    dtGridview_DLN1.Columns[colName].DefaultCellStyle.Format = "N2";
                    dtGridview_DLN1.Columns[colName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            if (dtGridview_DLN1.Columns.Contains("DocEntry"))
                dtGridview_DLN1.Columns["DocEntry"].Visible = false;
        }
        private void btnSearchAR_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCustomerAR.Text) && string.IsNullOrWhiteSpace(txtCustomerAR2.Text))
            {
                MessageBox.Show("Vui lòng nhập tên khách hàng hoặc mã khách hàng để tìm kiếm.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Company company = ConnectToSAP2();
            if (company == null)
                return;

            string cardName = txtCustomerAR.Text.Trim().Replace("'", "''");
            string cardCode = txtCustomerAR2.Text.Trim().Replace("'", "''");
            string fromDate = fromdatePickAR.Value.ToString("yyyy-MM-dd");
            string toDate = todatePickAR.Value.ToString("yyyy-MM-dd");

            string statusFilter = "";
            if (checkOpenAR.Checked && !checkClosedAR.Checked)
                statusFilter = "AND DocStatus = 'O'";
            else if (!checkOpenAR.Checked && checkClosedAR.Checked)
                statusFilter = "AND DocStatus = 'C'";

            string cardCodeFilter = string.IsNullOrWhiteSpace(cardCode) ? "" : $"CardCode LIKE '%{cardCode}%'";
            string cardNameFilter = string.IsNullOrWhiteSpace(cardName) ? "" : $"CardName LIKE N'%{cardName}%'";

            string whereFilter = "";
            if (!string.IsNullOrEmpty(cardCodeFilter) && !string.IsNullOrEmpty(cardNameFilter))
                whereFilter = $"({cardCodeFilter} AND {cardNameFilter})";
            else if (!string.IsNullOrEmpty(cardCodeFilter))
                whereFilter = cardCodeFilter;
            else if (!string.IsNullOrEmpty(cardNameFilter))
                whereFilter = cardNameFilter;

            string query = $@"
            SELECT ROW_NUMBER() OVER (ORDER BY DocEntry) AS STT,
            DocEntry, CardCode, CardName, DocStatus, DocType, 
            CONVERT(varchar(10), DocDate, 103) AS DocDate, 
            CONVERT(varchar(10), DocDueDate, 103) AS DocDueDate, 
            DocTotal
            FROM OINV WITH(NOLOCK)
            WHERE {whereFilter}
              AND DocDate BETWEEN '{fromDate}' AND '{toDate}'
              {statusFilter}
            ORDER BY DocEntry";

            Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            DataTable dt = new DataTable();
            dt.Columns.Add("STT");
            dt.Columns.Add("DocEntry");
            dt.Columns.Add("CardCode");
            dt.Columns.Add("CardName");
            dt.Columns.Add("DocStatus");
            dt.Columns.Add("DocType");
            dt.Columns.Add("DocDate");
            dt.Columns.Add("DocDueDate");
            dt.Columns.Add("DocTotal");

            while (!recordset.EoF)
            {
                DataRow row = dt.NewRow();
                row["STT"] = recordset.Fields.Item("STT").Value;
                row["DocEntry"] = recordset.Fields.Item("DocEntry").Value;
                row["CardCode"] = recordset.Fields.Item("CardCode").Value;
                row["CardName"] = recordset.Fields.Item("CardName").Value;
                row["DocStatus"] = recordset.Fields.Item("DocStatus").Value;
                row["DocType"] = recordset.Fields.Item("DocType").Value;
                row["DocDate"] = recordset.Fields.Item("DocDate").Value;
                row["DocDueDate"] = recordset.Fields.Item("DocDueDate").Value;
                row["DocTotal"] = recordset.Fields.Item("DocTotal").Value;
                dt.Rows.Add(row);
                recordset.MoveNext();
            }

            dtGridview_OINV.DataSource = dt;
            if (company.Connected)
                company.Disconnect();
        }

        private void dtGridview_OINV_SelectionChanged(object sender, EventArgs e)
        {
            if (dtGridview_OINV.SelectedRows.Count == 0)
                return;

            // Get DocEntry from the selected row
            var selectedRow = dtGridview_OINV.SelectedRows[0];
            if (selectedRow.Cells["DocEntry"].Value == null)
                return;

            int docEntry;
            if (!int.TryParse(selectedRow.Cells["DocEntry"].Value.ToString(), out docEntry))
                return;

            // Connect to SAP
            Company company = ConnectToSAP2();
            if (company == null)
                return;

            // Query INV1 lines for the selected DocEntry
            string query = $@"
            SELECT ROW_NUMBER() OVER (ORDER BY DocEntry) AS STT, 
            DocEntry, LineNum, ItemCode, Dscription, Quantity, Price, LineTotal, VatGroup, GTotal, WhsCode, AcctCode, BaseEntry
            FROM INV1 WITH(NOLOCK)
            WHERE DocEntry = {docEntry}
            ORDER BY LineNum";

            Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            // Load data into a DataTable
            DataTable dt = new DataTable();
            dt.Columns.Add("STT");
            dt.Columns.Add("DocEntry");
            dt.Columns.Add("LineNum");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("Dscription");
            dt.Columns.Add("Quantity");
            dt.Columns.Add("Price");
            dt.Columns.Add("LineTotal");
            dt.Columns.Add("VatGroup");
            dt.Columns.Add("GTotal");
            dt.Columns.Add("WhsCode");
            dt.Columns.Add("AcctCode");
            dt.Columns.Add("BaseEntry");

            while (!recordset.EoF)
            {
                DataRow row = dt.NewRow();
                row["STT"] = recordset.Fields.Item("STT").Value;
                row["DocEntry"] = recordset.Fields.Item("DocEntry").Value;
                row["LineNum"] = recordset.Fields.Item("LineNum").Value;
                row["ItemCode"] = recordset.Fields.Item("ItemCode").Value;
                row["Dscription"] = recordset.Fields.Item("Dscription").Value;
                row["Quantity"] = recordset.Fields.Item("Quantity").Value;
                row["Price"] = recordset.Fields.Item("Price").Value;
                row["LineTotal"] = recordset.Fields.Item("LineTotal").Value;
                row["VatGroup"] = recordset.Fields.Item("VatGroup").Value;
                row["GTotal"] = recordset.Fields.Item("GTotal").Value;
                row["WhsCode"] = recordset.Fields.Item("WhsCode").Value;
                row["AcctCode"] = recordset.Fields.Item("AcctCode").Value;
                row["BaseEntry"] = recordset.Fields.Item("BaseEntry").Value;
                dt.Rows.Add(row);
                recordset.MoveNext();
            }
            dtGridview_INV1.DataSource = dt;
            if (company.Connected)
                company.Disconnect();
        }
        private void btnImportODLN_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = new DataTable();

                    using (var workbook = new XLWorkbook(ofd.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var rows = worksheet.RangeUsed().RowsUsed().ToList();

                        if (rows.Count < 2)
                        {
                            MessageBox.Show("File phải có ít nhất 2 dòng (1 dòng tiêu đề nhóm, 1 dòng header).");
                            return;
                        }

                        // Lấy header từ dòng 2, bỏ qua cột trùng tên
                        var headerCells = rows[1].Cells().ToList();
                        List<int> validColumnIndexes = new List<int>();
                        for (int i = 0; i < headerCells.Count; i++)
                        {
                            string colName = headerCells[i].GetString();
                            if (!dt.Columns.Contains(colName) && !string.IsNullOrWhiteSpace(colName))
                            {
                                dt.Columns.Add(colName);
                                validColumnIndexes.Add(i);
                            }
                            // Nếu trùng tên thì bỏ qua index này
                        }
                        // Dữ liệu từ dòng 3 trở đi, chỉ lấy các cột hợp lệ
                        for (int i = 2; i < rows.Count; i++)
                        {
                            var dataCells = rows[i].Cells().ToList();
                            object[] rowData = new object[validColumnIndexes.Count];
                            for (int j = 0; j < validColumnIndexes.Count; j++)
                            {
                                int colIdx = validColumnIndexes[j];
                                object value = colIdx < dataCells.Count ? (object)dataCells[colIdx].Value.ToString() : DBNull.Value;
                                rowData[j] = value;
                            }
                            dt.Rows.Add(rowData);
                        }
                    }
                    dtGridview_ODLN.DataSource = dt;
                }
            }
        }
        private void btnSelectAllDeli_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dtGridview_ODLN.Rows)
            {
                row.Cells["Select"].Value = true;
            }
        }

        private void btnDeselectAllDeli_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dtGridview_ODLN.Rows)
            {
                row.Cells["Select"].Value = false;
            }
        }

        private void btnImportODLN_v2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = new DataTable();
                    try
                    {
                        // Thử đọc trực tiếp với FileShare
                        using (FileStream fileStream = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var workbook = new XLWorkbook(fileStream))
                        {
                            ProcessExcelData(workbook, dt);
                        }
                        dtGridview_ODLN_v2.DataSource = dt;
                    }
                    catch (IOException)
                    {
                        // Nếu vẫn lỗi, thử tạo file tạm thời
                        try
                        {
                            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(ofd.FileName));
                            File.Copy(ofd.FileName, tempFilePath, true);

                            using (var workbook = new XLWorkbook(tempFilePath))
                            {
                                ProcessExcelData(workbook, dt);
                            }

                            File.Delete(tempFilePath); // Xóa file tạm
                            dtGridview_ODLN_v2.DataSource = dt;
                        }
                        catch
                        {
                            MessageBox.Show("Không thể import file Excel. Vui lòng thử lại.",
                                          "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi import file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ProcessExcelData(XLWorkbook workbook, DataTable dt)
        {
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed().ToList();

            if (rows.Count < 2)
            {
                MessageBox.Show("File phải có ít nhất 2 dòng (1 dòng tiêu đề nhóm, 1 dòng header).");
                return;
            }

            var headerCells = rows[1].Cells().ToList();
            List<int> validColumnIndexes = new List<int>();
            for (int i = 0; i < headerCells.Count; i++)
            {
                string colName = headerCells[i].GetString();
                if (!dt.Columns.Contains(colName) && !string.IsNullOrWhiteSpace(colName))
                {
                    dt.Columns.Add(colName);
                    validColumnIndexes.Add(i);
                }
            }

            for (int i = 2; i < rows.Count; i++)
            {
                var dataCells = rows[i].Cells().ToList();
                object[] rowData = new object[validColumnIndexes.Count];
                for (int j = 0; j < validColumnIndexes.Count; j++)
                {
                    int colIdx = validColumnIndexes[j];
                    object value = colIdx < dataCells.Count ? (object)dataCells[colIdx].Value.ToString() : DBNull.Value;
                    rowData[j] = value;
                }
                dt.Rows.Add(rowData);
            }
        }
        private void btnGenARExcel_v2_Click(object sender, EventArgs e)
        {
            if (dtGridview_ODLN_v2.DataSource == null)
            {
                MessageBox.Show("Vui lòng import dữ liệu ODLN trước khi tạo AR Invoice.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DataTable dtODLN = dtGridview_ODLN_v2.DataSource as DataTable;
            if (dtODLN.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu để tạo AR Invoice.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Company company = ConnectToSAP2();
            if (company == null)
                return;

            // Decrypt UserID from App.config
            string decryptKey = "34bcf4830ab7dfa70e9fd4c5daacd7ed2983099b31d65c8c4089d3d2b2b26b40";
            string decryptedUserID = DecryptAES(ConfigurationManager.AppSettings["UserID"], decryptKey);

            // Gom kết quả
            List<string> successList = new List<string>();
            List<string> errorList = new List<string>();

            // Group các dòng theo DocEntry
            var docEntryGroups = dtODLN.AsEnumerable()
                .GroupBy(r => r["DocEntry"].ToString());

            foreach (var group in docEntryGroups)
            {
                DataRow firstRow = group.First();
                Documents arInvoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);

                // Header
                if (dtODLN.Columns.Contains("CardCode"))
                    arInvoice.CardCode = firstRow["CardCode"].ToString();

                if (dtODLN.Columns.Contains("DocDate"))
                {
                    var docDateStr = firstRow["DocDate"]?.ToString();
                    DateTime docDate;
                    if (DateTime.TryParse(docDateStr, out docDate))
                        arInvoice.DocDate = docDate;
                    else if (DateTime.TryParseExact(docDateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDate))
                        arInvoice.DocDate = docDate;
                }

                if (dtODLN.Columns.Contains("DocDueDate"))
                {
                    var docDueDateStr = firstRow["DocDueDate"]?.ToString();
                    DateTime docDueDate;
                    if (DateTime.TryParse(docDueDateStr, out docDueDate))
                        arInvoice.DocDueDate = docDueDate;
                    else if (DateTime.TryParseExact(docDueDateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out docDueDate))
                        arInvoice.DocDueDate = docDueDate;
                }
                if (dtODLN.Columns.Contains("U_InvCode"))
                    arInvoice.UserFields.Fields.Item("U_InvCode").Value = firstRow["U_InvCode"].ToString();
                if (dtODLN.Columns.Contains("U_InvSerial"))
                    arInvoice.UserFields.Fields.Item("U_InvSerial").Value = firstRow["U_InvSerial"].ToString();
                if (dtODLN.Columns.Contains("U_DeclarePd"))
                {
                    var declarePdStr = firstRow["U_DeclarePd"]?.ToString();
                    DateTime declarePd;
                    if (DateTime.TryParse(declarePdStr, out declarePd))
                        arInvoice.UserFields.Fields.Item("U_DeclarePd").Value = declarePd;
                    else if (DateTime.TryParseExact(declarePdStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out declarePd))
                        arInvoice.UserFields.Fields.Item("U_DeclarePd").Value = declarePd;
                }
                if (dtODLN.Columns.Contains("U_InvDate"))
                {
                    var U_InvDateStr = firstRow["U_InvDate"]?.ToString();
                    DateTime U_InvDate;
                    if (DateTime.TryParse(U_InvDateStr, out U_InvDate))
                        arInvoice.UserFields.Fields.Item("U_InvDate").Value = U_InvDate;
                    else if (DateTime.TryParseExact(U_InvDateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out U_InvDate))
                        arInvoice.UserFields.Fields.Item("U_InvDate").Value = U_InvDate;
                }
                if (dtODLN.Columns.Contains("U_DescVN"))
                    arInvoice.UserFields.Fields.Item("U_DescVN").Value = firstRow["U_DescVN"].ToString();
                //arInvoice.UserFields.Fields.Item("U_MailLog").Value = $"From Winform, created by {decryptedUserID} on {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                // Body
                bool firstLine = true;
                foreach (DataRow lineRow in group)
                {
                    if (!firstLine)
                        arInvoice.Lines.Add();
                    if (dtODLN.Columns.Contains("ItemCode"))
                        arInvoice.Lines.ItemCode = lineRow["ItemCode"].ToString();
                    if (dtODLN.Columns.Contains("Quantity") && !string.IsNullOrWhiteSpace(lineRow["Quantity"]?.ToString()))
                        arInvoice.Lines.Quantity = Convert.ToDouble(lineRow["Quantity"]);
                    if (dtODLN.Columns.Contains("Price") && !string.IsNullOrWhiteSpace(lineRow["Price"]?.ToString()))
                        arInvoice.Lines.Price = Convert.ToDouble(lineRow["Price"]);
                    if (dtODLN.Columns.Contains("WhsCode"))
                        arInvoice.Lines.WarehouseCode = lineRow["WhsCode"].ToString();
                    if (dtODLN.Columns.Contains("AcctCode"))
                        arInvoice.Lines.AccountCode = lineRow["AcctCode"].ToString();
                    if (dtODLN.Columns.Contains("VatGroup"))
                        arInvoice.Lines.VatGroup = lineRow["VatGroup"].ToString();
                    arInvoice.Lines.BaseType = 15;
                    if (dtODLN.Columns.Contains("BaseEntry") && !string.IsNullOrWhiteSpace(lineRow["BaseEntry"]?.ToString()))
                        arInvoice.Lines.BaseEntry = Convert.ToInt32(lineRow["BaseEntry"]);
                    if (dtODLN.Columns.Contains("BaseLine") && !string.IsNullOrWhiteSpace(lineRow["BaseLine"]?.ToString()))
                        arInvoice.Lines.BaseLine = Convert.ToInt32(lineRow["BaseLine"]);
                    //arInvoice.Lines.UserFields.Fields.Item("U_Remarks").Value = $"From Winform, created by {decryptedUserID} (Windows user: {Environment.UserDomainName}/{Environment.UserName}) on {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                    firstLine = false;
                }

                int result = arInvoice.Add();
                if (result == 0)
                {
                    successList.Add(group.Key); // lưu lại docentry thành công
                }
                else
                {
                    string errMsg;
                    int errCode;
                    company.GetLastError(out errCode, out errMsg);
                    errorList.Add($"{group.Key} (ErrCode {errCode}: {errMsg})");
                }
            }

            // Sau khi chạy xong tất cả mới hiện 1 thông báo tổng hợp
            string message = "";
            if (successList.Count > 0)
                message += $"Tạo AR Invoice thành công cho DocEntry: {string.Join(", ", successList)}.\n";
            if (errorList.Count > 0)
                message += $"Không tạo được cho DocEntry: {string.Join("; ", errorList)}.";

            MessageBox.Show(message, "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (company.Connected)
                company.Disconnect();
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage4)
            {
                this.MaximizeBox = false;
            }
            else
            {
                this.MaximizeBox = true;
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void txtCustomerAR_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSearchAR_Click(btnSearchAR, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void txtCustomer_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSearch_Click(btn4PsSearch, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnLogin_Click(btnLogin, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void btnRefresh_Click(object sender, EventArgs e)
        {
            dtGridview_ODLN_v2.DataSource = null;
        }
        private void txtCustomerAR2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSearchAR_Click(btn4PsSearch, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void txtCustomer2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSearch_Click(btn4PsSearch, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        private void CenterLabel1InForm()
        {
            if (label1 != null)
            {
                int yPosition = -1;
                int x = (this.ClientSize.Width - label1.Width) / 2;
                int y = yPosition == -1 ? label1.Location.Y : yPosition;
                label1.Location = new Point(Math.Max(0, x), y);
            }
        }
        private void ARInvoice_Resize(object sender, EventArgs e)
        {
            CenterLabel1InForm();
            CenterLabelInTabPage();
        }
        private void CenterControlInTabPage(Control control, int yPosition = -1, Control referenceControl = null)
        {
            int x;
            if (referenceControl != null)
            {
                x = referenceControl.Location.X;
            }
            else
            {
                x = (this.ClientSize.Width - control.Width) / 2;
            }

            int y = yPosition == -1 ? control.Location.Y : yPosition;
            control.Location = new Point(Math.Max(0, x), y);
        }
        private void CenterLabelInTabPage()
        {
            CenterControlInTabPage(txtUserID);
            CenterControlInTabPage(txtPassword);
            CenterControlInTabPage(label10);
            CenterControlInTabPage(label13, -1, txtUserID);
            CenterControlInTabPage(label14, -1, txtUserID);
            CenterControlInTabPage(btnLogin);
            CenterControlInTabPage(btnCancel);
        }
        private void btnExportDeli_Click(object sender, EventArgs e)
        {
            var selectedDocEntries = new List<int>();
            foreach (DataGridViewRow row in dtGridview_ODLN.Rows)
            {
                if (row.IsNewRow) continue;
                var cell = row.Cells["Select"] as DataGridViewCheckBoxCell;
                if (cell != null && cell.Value != null && (bool)cell.Value)
                {
                    if (int.TryParse(row.Cells["DocEntry"].Value?.ToString(), out int docEntry))
                        selectedDocEntries.Add(docEntry);
                }
            }
            if (selectedDocEntries.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một Delivery để export.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Company company = ConnectToSAP2();
            if (company == null)
                return;

            DataTable dtExport = new DataTable();
            dtExport.Columns.Add("DocDate");
            dtExport.Columns.Add("DocDueDate");
            dtExport.Columns.Add("CardCode");
            dtExport.Columns.Add("ItemCode");
            dtExport.Columns.Add("Quantity");
            dtExport.Columns.Add("Price");
            dtExport.Columns.Add("WhsCode");
            dtExport.Columns.Add("GLAccount");
            dtExport.Columns.Add("VatGroup");
            dtExport.Columns.Add("BaseEntry");
            dtExport.Columns.Add("BaseLine");

            foreach (int docEntry in selectedDocEntries)
            {
                string query = $@"
                SELECT 
                    CONVERT(VARCHAR(8), GETDATE(), 112) AS DocDate,
                    CONVERT(VARCHAR(8), GETDATE(), 112) AS DocDueDate,
                    T1.CardCode,
                    T2.ItemCode,
                    (T2.Quantity - ISNULL(SUM(R1.Quantity),0)) AS Quantity,   -- số lượng thực tế
                    T2.Price,
                    T2.WhsCode,
                    T2.AcctCode as GLAccount,
                    T2.VatGroup,
                    T2.DocEntry as BaseEntry,
                    T2.LineNum AS BaseLine
                FROM ODLN T1 WITH(NOLOCK)
                JOIN DLN1 T2 WITH(NOLOCK) ON T1.DocEntry = T2.DocEntry
                LEFT JOIN RDN1 R1 WITH(NOLOCK) ON R1.BaseEntry = T2.DocEntry AND R1.BaseLine = T2.LineNum
                WHERE T1.DocEntry = {docEntry} AND T2.LineStatus LIKE '%O%'
                GROUP BY
                    T1.DocDate, T1.DocDueDate, T1.CardCode,
                    T2.ItemCode, T2.Quantity, T2.Price,
                    T2.WhsCode, T2.AcctCode, T2.VatGroup,
                    T2.DocEntry, T2.LineNum";

                Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rs.DoQuery(query);

                while (!rs.EoF)
                {
                    DataRow dr = dtExport.NewRow();
                    dr["DocDate"] = rs.Fields.Item("DocDate").Value;
                    dr["DocDueDate"] = rs.Fields.Item("DocDueDate").Value;
                    dr["CardCode"] = rs.Fields.Item("CardCode").Value;
                    dr["ItemCode"] = rs.Fields.Item("ItemCode").Value;
                    dr["Quantity"] = rs.Fields.Item("Quantity").Value;
                    dr["Price"] = rs.Fields.Item("Price").Value;
                    dr["WhsCode"] = rs.Fields.Item("WhsCode").Value;
                    dr["GLAccount"] = rs.Fields.Item("GLAccount").Value;
                    dr["VatGroup"] = rs.Fields.Item("VatGroup").Value;
                    dr["BaseEntry"] = rs.Fields.Item("BaseEntry").Value;
                    dr["BaseLine"] = rs.Fields.Item("BaseLine").Value;
                    dtExport.Rows.Add(dr);
                    rs.MoveNext();
                }
            }
            if (company.Connected)
                company.Disconnect();
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx";
                sfd.FileName = "DeliveryExport.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Delivery");
                        ws.Cell(1, 1).InsertTable(dtExport, "Delivery", true);
                        wb.SaveAs(sfd.FileName);
                    }
                    MessageBox.Show("Export thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            if (btnLogout != null && btnLogout.Visible)
            {
                btnLogout.Visible = false;
                if (btnLogout.Parent != null)
                {
                    btnLogout.Parent.Controls.Remove(btnLogout);
                }
            }
            this.Size = new Size(444, 500);
            if (tabControl1.TabPages.Count > 0)
            {
                tabControl1.TabPages.Clear();
            }
            if (tabPage4 != null && !tabControl1.TabPages.Contains(tabPage4))
            {
                tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tabPage4);
            }
        }
        private void btnBkavCheckConnect_Click(object sender, EventArgs e)
        {
            using (var conn = ConnectToSQL())
            {
                if (conn == null)
                {
                    MessageBox.Show("Kết nối SQL Server thất bại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                try
                {
                    MessageBox.Show("Kết nối SQL Server thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }
        }
        private HashSet<int> modifiedRows = new HashSet<int>();

        private void dtGridview_BKAVCus_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex > 0)) // Không track thay đổi cột Select
            {
                modifiedRows.Add(e.RowIndex);
            }
        }
        private void txt4PsSearch_TextChanged(object sender, EventArgs e)
        {
            if (dtHeader == null || dtHeader.Rows.Count == 0) return;

            string searchValue = txt4PsSearch.Text.Trim().Replace("'", "''");

            if (string.IsNullOrEmpty(searchValue))
            {
                bsHeader.RemoveFilter();
            }
            else
            {
                // Bạn có thể tùy chỉnh lọc theo nhiều cột cùng lúc (CardCode, CardName, DocEntry...)
                // Sử dụng Convert(..., 'System.String') để đảm bảo các cột số cũng có thể lọc kiểu LIKE
                bsHeader.Filter = string.Format(
                    "(Convert(DocEntry, 'System.String') LIKE '%{0}%') OR " +
                    "(CardCode LIKE '%{0}%') OR " +
                    "(CardName LIKE '%{0}%')",
                    searchValue);
            }
        }

        private void btnKitSearch_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Kiểm tra cột Checkbox "Chọn"
                if (!KitdtGridview_ODLN.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn selectColumn = new DataGridViewCheckBoxColumn();
                    selectColumn.Name = "Select";
                    selectColumn.HeaderText = "Chọn";
                    selectColumn.Width = 65;
                    KitdtGridview_ODLN.Columns.Insert(0, selectColumn);
                }

                // 2. Kiểm tra điều kiện đầu vào
                if (string.IsNullOrWhiteSpace(KittxtCustomer.Text) && string.IsNullOrWhiteSpace(KittxtCustomer2.Text))
                {
                    MessageBox.Show("Vui lòng nhập tên khách hàng hoặc mã khách hàng để tìm kiếm.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 3. Xóa Cache cũ trước khi nạp dữ liệu mới
                _invoiceLinesCache.Clear();

                // 4. Kết nối SAP
                Company company = ConnectToSAP3();
                if (company == null) return;

                // 5. Chuẩn bị biến lọc dữ liệu
                string cardName = KittxtCustomer.Text.Trim().Replace("'", "''");
                string cardCode = KittxtCustomer2.Text.Trim().Replace("'", "''");
                string fromDate = KitfromdatePick.Value.ToString("yyyy-MM-dd");
                string toDate = KittodatePick.Value.ToString("yyyy-MM-dd");

                string statusFilter = "";
                if (KitcheckBox_Open.Checked && !KitcheckBox_Closed.Checked)
                    statusFilter = "AND DocStatus = 'O'";
                else if (!KitcheckBox_Open.Checked && KitcheckBox_Closed.Checked)
                    statusFilter = "AND DocStatus = 'C'";

                string cardCodeFilter = string.IsNullOrWhiteSpace(cardCode) ? "" : $"T0.CardCode LIKE '%{cardCode}%'";
                string cardNameFilter = string.IsNullOrWhiteSpace(cardName) ? "" : $"T0.CardName LIKE N'%{cardName}%'";

                string whereFilter = "";
                if (!string.IsNullOrEmpty(cardCodeFilter) && !string.IsNullOrEmpty(cardNameFilter))
                    whereFilter = $"({cardCodeFilter} AND {cardNameFilter})";
                else if (!string.IsNullOrEmpty(cardCodeFilter))
                    whereFilter = cardCodeFilter;
                else if (!string.IsNullOrEmpty(cardNameFilter))
                    whereFilter = cardNameFilter;

                try
                {
                    string queryHeader = $@"
                    SELECT 
                        ROW_NUMBER() OVER (ORDER BY T0.DocEntry) AS STT,
                        T0.DocEntry, T0.CardCode, T0.CardName, T0.DocStatus, T0.DocType, 
                        CONVERT(varchar(10), T0.DocDate, 103) AS DocDate, 
                        CONVERT(varchar(10), T0.DocDueDate, 103) AS DocDueDate,  
                        T0.DocTotal, T0.SlpCode, S.Active,
                        ISNULL(R0.DocEntry, 0) AS ReturnDocEntry
                    FROM ODLN T0 WITH(NOLOCK)
                    LEFT JOIN OCRD C WITH(NOLOCK) ON C.CardCode = T0.CardCode
                    LEFT JOIN OSLP S WITH(NOLOCK) ON S.SlpCode = T0.SlpCode
                    OUTER APPLY (
                        SELECT TOP 1 R0.DocEntry
                        FROM ORDN R0 WITH(NOLOCK)
                        INNER JOIN RDN1 R1 ON R1.DocEntry = R0.DocEntry
                        WHERE R1.BaseEntry = T0.DocEntry AND R0.CANCELED = 'N'
                        ORDER BY R0.DocEntry
                    ) R0
                    WHERE {whereFilter} 
                      AND T0.DocDate BETWEEN '{fromDate}' AND '{toDate}' 
                      AND ISNULL(C.FrozenFor, 'N') <> 'Y'
                      {statusFilter}
                    ORDER BY T0.DocEntry";

                    Recordset rsHeader = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsHeader.DoQuery(queryHeader);

                    // Nạp dữ liệu vào DataTable cấp Class
                    dtHeader = ConvertRecordsetToDataTable(rsHeader);

                    // Gán DataTable vào BindingSource
                    bsHeader.DataSource = dtHeader;

                    // Gán BindingSource vào DataGridView
                    KitdtGridview_ODLN.DataSource = bsHeader;

                    // --- BƯỚC 7: QUERY TOÀN BỘ CHI TIẾT (LINES) VÀO CACHE ---
                    // Sử dụng INNER JOIN để lọc các Lines thuộc về các Headers ở trên
                    string queryAllLines = $@"
                    WITH ReturnSummary AS (
                        SELECT R1.BaseEntry, R1.BaseLine, SUM(R1.Quantity) AS ReturnQty
                        FROM RDN1 R1 WITH(NOLOCK)
                        INNER JOIN ORDN R0 WITH(NOLOCK) ON R0.DocEntry = R1.DocEntry
                        WHERE R0.CANCELED = 'N'
                        GROUP BY R1.BaseEntry, R1.BaseLine
                    ),
                    InvoiceSummary AS (
                        SELECT I1.BaseEntry, I1.BaseLine, SUM(I1.Quantity) AS InvoicedQty, MAX(I1.DocEntry) AS InvDocEntry
                        FROM INV1 I1 WITH(NOLOCK)
                        INNER JOIN OINV I0 WITH(NOLOCK) ON I0.DocEntry = I1.DocEntry
                        WHERE I0.CANCELED = 'N' AND I1.BaseType = 15
                        GROUP BY I1.BaseEntry, I1.BaseLine
                    )
                    SELECT 
                        T0.DocEntry, T0.LineNum, T0.ItemCode, T0.Dscription, T0.Quantity, 
                        ISNULL(TR.ReturnQty, 0) AS ReturnQty, 
                        ISNULL(TI.InvoicedQty, 0) AS InvoicedQty,
                        (T0.Quantity - ISNULL(TR.ReturnQty, 0) - ISNULL(TI.InvoicedQty, 0)) AS NetQuantity,
                        T0.Price, T0.LineTotal, T0.VatGroup, T0.GTotal, T0.WhsCode, 
                        T0.AcctCode, T0.LineStatus, T0.TrgetEntry,
                        ISNULL(TI.InvDocEntry, 0) AS InvDocEntry
                    FROM DLN1 T0 WITH(NOLOCK)
                    INNER JOIN ODLN H WITH(NOLOCK) ON T0.DocEntry = H.DocEntry
                    LEFT JOIN ReturnSummary TR ON TR.BaseEntry = T0.DocEntry AND TR.BaseLine = T0.LineNum
                    LEFT JOIN InvoiceSummary TI ON TI.BaseEntry = T0.DocEntry AND TI.BaseLine = T0.LineNum
                    WHERE {whereFilter.Replace("T0.CardCode", "H.CardCode").Replace("T0.CardName", "H.CardName")} 
                      AND H.DocDate BETWEEN '{fromDate}' AND '{toDate}' 
                      {statusFilter.Replace("DocStatus", "H.DocStatus")}";

                    Recordset rsLines = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    rsLines.DoQuery(queryAllLines);
                    DataTable dtAllLines = ConvertRecordsetToDataTable(rsLines);

                    // --- BƯỚC 8: PHÂN TÁCH DỮ LIỆU VÀO DICTIONARY (CACHE) ---
                    if (dtAllLines.Rows.Count > 0)
                    {
                        // Group dữ liệu theo DocEntry bằng LINQ
                        var grouped = dtAllLines.AsEnumerable().GroupBy(r => Convert.ToInt32(r["DocEntry"]));
                        foreach (var group in grouped)
                        {
                            _invoiceLinesCache[group.Key] = group.CopyToDataTable();
                        }
                    }

                    MessageBox.Show($"Tìm thấy {KitdtGridview_ODLN.Rows.Count} hóa đơn.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    if (company.Connected) company.Disconnect();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void KitFormatDetailGridView()
        {
            if (KitdtGridview_DLN1.Columns.Count == 0) return;

            Dictionary<string, string> headers = new Dictionary<string, string>
            {
                { "Dscription", "Item Name" },
                { "AcctCode", "G/L Account" },
                { "TrgetEntry", "Return DocEntry" },
                { "InvDocEntry", "Invoice DocEntry" }, // Cột mới thêm
                { "InvoicedQty", "Invoiced Qty" },     // Cột mới thêm
                { "VatGroup", "Tax Code" },
                { "GTotal", "Gross Total" },
                { "NetQuantity", "Remaining Qty" }      // Đổi tên cho rõ nghĩa
            };

            foreach (var header in headers)
            {
                if (KitdtGridview_DLN1.Columns.Contains(header.Key))
                {
                    KitdtGridview_DLN1.Columns[header.Key].HeaderText = header.Value;
                }
            }

            // Cập nhật các cột số lượng cần định dạng
            string[] numericColumns = { "Quantity", "ReturnQty", "InvoicedQty", "NetQuantity", "Price", "LineTotal", "GTotal" };
            foreach (string colName in numericColumns)
            {
                if (KitdtGridview_DLN1.Columns.Contains(colName))
                {
                    KitdtGridview_DLN1.Columns[colName].DefaultCellStyle.Format = "N2";
                    KitdtGridview_DLN1.Columns[colName].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }

            if (KitdtGridview_DLN1.Columns.Contains("DocEntry"))
                KitdtGridview_DLN1.Columns["DocEntry"].Visible = false;
        }

        private void KitdtGridview_ODLN_SelectionChanged(object sender, EventArgs e)
        {
            if (KitdtGridview_ODLN.SelectedRows.Count == 0) return;
            var selectedRow = KitdtGridview_ODLN.SelectedRows[0];
            if (selectedRow.Cells["DocEntry"].Value == null) return;
            int docEntry = Convert.ToInt32(selectedRow.Cells["DocEntry"].Value);
            if (_invoiceLinesCache.ContainsKey(docEntry))
            {
                KitdtGridview_DLN1.DataSource = _invoiceLinesCache[docEntry];
            }
            else
            {
                // Trường hợp hóa đơn không có dòng nào (hiếm gặp)
                KitdtGridview_DLN1.DataSource = null;
            }
            FormatDetailGridView();
        }

        private void KitbtnCheck_connection_Click(object sender, EventArgs e)
        {
            ConnectToSAP3();
        }

        private void KittxtSearch_TextChanged(object sender, EventArgs e)
        {
            if (dtHeader == null || dtHeader.Rows.Count == 0) return;

            string searchValue = KittxtSearch.Text.Trim().Replace("'", "''");

            if (string.IsNullOrEmpty(searchValue))
            {
                bsHeader.RemoveFilter();
            }
            else
            {
                bsHeader.Filter = string.Format(
                    "(Convert(DocEntry, 'System.String') LIKE '%{0}%') OR " +
                    "(CardCode LIKE '%{0}%') OR " +
                    "(CardName LIKE '%{0}%')",
                    searchValue);
            }
        }

        private async void KitbtnCopyToInvoice_Click(object sender, EventArgs e)
        {
            // BƯỚC 1: LẤY DỮ LIỆU TỪ UI (Phải làm ở UI Thread)
            var selectedRowsData = KitdtGridview_ODLN.Rows.Cast<DataGridViewRow>()
                .Where(row => row.Cells["Select"].Value != null && (bool)row.Cells["Select"].Value == true)
                .Select(row => new {
                    CardCode = row.Cells["CardCode"].Value.ToString(),
                    DocEntry = Convert.ToInt32(row.Cells["DocEntry"].Value),
                    SlpCode = Convert.ToInt32(row.Cells["SlpCode"]?.Value ?? -1),
                    Active = row.Cells["Active"]?.Value?.ToString()
                })
                .ToList();

            if (selectedRowsData.Count == 0)
            {
                MessageBox.Show("Vui lòng chọn ít nhất một phiếu Delivery.");
                return;
            }

            // BƯỚC 2: CHUẨN BỊ UI
            this.Cursor = Cursors.WaitCursor;
            KitbtnCopyToInvoice.Enabled = false;

            Company company = ConnectToSAP3(); // Kết nối SAP
            if (company == null)
            {
                this.Cursor = Cursors.Default;
                KitbtnCopyToInvoice.Enabled = true;
                return;
            }

            // BƯỚC 3: XỬ LÝ NẶNG Ở BACKGROUND
            var resultSummary = await Task.Run(() =>
            {
                var successList = new List<string>();
                var errorList = new List<string>();
                var successDocEntries = new List<int>(); // Thêm list này để theo dõi ID thành công

                var groups = selectedRowsData.GroupBy(x => x.CardCode);

                foreach (var group in groups)
                {
                    string cardCode = group.Key;
                    List<int> docEntries = group.Select(x => x.DocEntry).ToList();
                    int finalSlp = group.Any(x => x.Active == "Y") ? group.First().SlpCode : -1;

                    var res = CreateARInvoiceInternal(company, cardCode, docEntries, finalSlp);

                    if (res.Success)
                    {
                        successList.Add($"{cardCode} -> {res.NewDocEntry}");
                        successDocEntries.AddRange(docEntries); // Lưu lại các ID đã xong
                    }
                    else
                    {
                        errorList.Add($"{cardCode}: {res.ErrorMessage}");
                    }
                }

                return new { successList, errorList, successDocEntries };
            });

            // BƯỚC 4: KẾT THÚC VÀ HIỂN THỊ KẾT QUẢ
            if (company.Connected) company.Disconnect();
            this.Cursor = Cursors.Default;
            KitbtnCopyToInvoice.Enabled = true;
            KitRefreshGridViewAfterSuccess(resultSummary.successDocEntries);
            ShowFinalReport(resultSummary.successList, resultSummary.errorList);
        }

        private void KitbtnSearchAR_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(KittxtCustomerAR.Text) && string.IsNullOrWhiteSpace(KittxtCustomerAR2.Text))
            {
                MessageBox.Show("Vui lòng nhập tên khách hàng hoặc mã khách hàng để tìm kiếm.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Company company = ConnectToSAP3();
            if (company == null)
                return;

            string cardName = KittxtCustomerAR.Text.Trim().Replace("'", "''");
            string cardCode = KittxtCustomerAR2.Text.Trim().Replace("'", "''");
            string fromDate = KitfromdatePickAR.Value.ToString("yyyy-MM-dd");
            string toDate = KittodatePickAR.Value.ToString("yyyy-MM-dd");

            string statusFilter = "";
            if (KitcheckOpenAR.Checked && !KitcheckClosedAR.Checked)
                statusFilter = "AND DocStatus = 'O'";
            else if (!KitcheckOpenAR.Checked && KitcheckClosedAR.Checked)
                statusFilter = "AND DocStatus = 'C'";

            string cardCodeFilter = string.IsNullOrWhiteSpace(cardCode) ? "" : $"CardCode LIKE '%{cardCode}%'";
            string cardNameFilter = string.IsNullOrWhiteSpace(cardName) ? "" : $"CardName LIKE N'%{cardName}%'";

            string whereFilter = "";
            if (!string.IsNullOrEmpty(cardCodeFilter) && !string.IsNullOrEmpty(cardNameFilter))
                whereFilter = $"({cardCodeFilter} AND {cardNameFilter})";
            else if (!string.IsNullOrEmpty(cardCodeFilter))
                whereFilter = cardCodeFilter;
            else if (!string.IsNullOrEmpty(cardNameFilter))
                whereFilter = cardNameFilter;

            string query = $@"
            SELECT ROW_NUMBER() OVER (ORDER BY DocEntry) AS STT,
            DocEntry, CardCode, CardName, DocStatus, DocType, 
            CONVERT(varchar(10), DocDate, 103) AS DocDate, 
            CONVERT(varchar(10), DocDueDate, 103) AS DocDueDate, 
            DocTotal
            FROM OINV WITH(NOLOCK)
            WHERE {whereFilter}
              AND DocDate BETWEEN '{fromDate}' AND '{toDate}'
              {statusFilter}
            ORDER BY DocEntry";

            Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            DataTable dt = new DataTable();
            dt.Columns.Add("STT");
            dt.Columns.Add("DocEntry");
            dt.Columns.Add("CardCode");
            dt.Columns.Add("CardName");
            dt.Columns.Add("DocStatus");
            dt.Columns.Add("DocType");
            dt.Columns.Add("DocDate");
            dt.Columns.Add("DocDueDate");
            dt.Columns.Add("DocTotal");

            while (!recordset.EoF)
            {
                DataRow row = dt.NewRow();
                row["STT"] = recordset.Fields.Item("STT").Value;
                row["DocEntry"] = recordset.Fields.Item("DocEntry").Value;
                row["CardCode"] = recordset.Fields.Item("CardCode").Value;
                row["CardName"] = recordset.Fields.Item("CardName").Value;
                row["DocStatus"] = recordset.Fields.Item("DocStatus").Value;
                row["DocType"] = recordset.Fields.Item("DocType").Value;
                row["DocDate"] = recordset.Fields.Item("DocDate").Value;
                row["DocDueDate"] = recordset.Fields.Item("DocDueDate").Value;
                row["DocTotal"] = recordset.Fields.Item("DocTotal").Value;
                dt.Rows.Add(row);
                recordset.MoveNext();
            }

            KitdtGridview_OINV.DataSource = dt;
            if (company.Connected)
                company.Disconnect();
        }

        private void KitdtGridview_OINV_SelectionChanged(object sender, EventArgs e)
        {
            if (KitdtGridview_OINV.SelectedRows.Count == 0)
                return;

            // Get DocEntry from the selected row
            var selectedRow = KitdtGridview_OINV.SelectedRows[0];
            if (selectedRow.Cells["DocEntry"].Value == null)
                return;

            int docEntry;
            if (!int.TryParse(selectedRow.Cells["DocEntry"].Value.ToString(), out docEntry))
                return;

            // Connect to SAP
            Company company = ConnectToSAP3();
            if (company == null)
                return;

            // Query INV1 lines for the selected DocEntry
            string query = $@"
            SELECT ROW_NUMBER() OVER (ORDER BY DocEntry) AS STT, 
            DocEntry, LineNum, ItemCode, Dscription, Quantity, Price, LineTotal, VatGroup, GTotal, WhsCode, AcctCode, BaseEntry
            FROM INV1 WITH(NOLOCK)
            WHERE DocEntry = {docEntry}
            ORDER BY LineNum";

            Recordset recordset = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(query);

            // Load data into a DataTable
            DataTable dt = new DataTable();
            dt.Columns.Add("STT");
            dt.Columns.Add("DocEntry");
            dt.Columns.Add("LineNum");
            dt.Columns.Add("ItemCode");
            dt.Columns.Add("Dscription");
            dt.Columns.Add("Quantity");
            dt.Columns.Add("Price");
            dt.Columns.Add("LineTotal");
            dt.Columns.Add("VatGroup");
            dt.Columns.Add("GTotal");
            dt.Columns.Add("WhsCode");
            dt.Columns.Add("AcctCode");
            dt.Columns.Add("BaseEntry");

            while (!recordset.EoF)
            {
                DataRow row = dt.NewRow();
                row["STT"] = recordset.Fields.Item("STT").Value;
                row["DocEntry"] = recordset.Fields.Item("DocEntry").Value;
                row["LineNum"] = recordset.Fields.Item("LineNum").Value;
                row["ItemCode"] = recordset.Fields.Item("ItemCode").Value;
                row["Dscription"] = recordset.Fields.Item("Dscription").Value;
                row["Quantity"] = recordset.Fields.Item("Quantity").Value;
                row["Price"] = recordset.Fields.Item("Price").Value;
                row["LineTotal"] = recordset.Fields.Item("LineTotal").Value;
                row["VatGroup"] = recordset.Fields.Item("VatGroup").Value;
                row["GTotal"] = recordset.Fields.Item("GTotal").Value;
                row["WhsCode"] = recordset.Fields.Item("WhsCode").Value;
                row["AcctCode"] = recordset.Fields.Item("AcctCode").Value;
                row["BaseEntry"] = recordset.Fields.Item("BaseEntry").Value;
                dt.Rows.Add(row);
                recordset.MoveNext();
            }
            KitdtGridview_INV1.DataSource = dt;
            if (company.Connected)
                company.Disconnect();
        }
    }
}