using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;

namespace AutoGifteeBox
{
    public class ChromeDriverUpdater
    {
        public static void UpdateChromeDriver()
        {
            try
            {
                string chromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
                var versionInfo = FileVersionInfo.GetVersionInfo(chromePath);
                string chromeVersion = versionInfo.ProductVersion;
                string majorVersion = chromeVersion.Split('.')[0];

                Console.WriteLine($"Chrome version: {chromeVersion}");

                using (var client = new WebClient())
                {
                    string driverUrl = $"https://storage.googleapis.com/chrome-for-testing-public/{chromeVersion}/win64/chromedriver-win64.zip";
                    string tempZipPath = Path.Combine(Path.GetTempPath(), "chromedriver.zip");
                    client.DownloadFile(driverUrl, tempZipPath);
                    Console.WriteLine("Downloaded ChromeDriver zip.");

                    string targetPath = AppDomain.CurrentDomain.BaseDirectory;
                    using (ZipArchive archive = ZipFile.OpenRead(tempZipPath))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith("chromedriver.exe", StringComparison.OrdinalIgnoreCase))
                            {
                                string destinationPath = Path.Combine(targetPath, "chromedriver.exe");
                                entry.ExtractToFile(destinationPath, overwrite: true);
                                Console.WriteLine($"Extracted to: {destinationPath}");
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating ChromeDriver: {ex.Message}");
            }
        }
    }
}
