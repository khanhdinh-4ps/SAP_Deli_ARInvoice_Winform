using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;

internal class Program
{
    static int Main(string[] args)
    {
        try
        {
            if (args.Length < 2)
            {
                Log("Thiếu tham số!");
                return 1;
            }

            string zipFile = args[0];
            string mainExe = args[1];
            string appFolder = Path.GetDirectoryName(mainExe);

            Log($"Updater started. zipFile={zipFile}, mainExe={mainExe}");

            // Wait for the main application to exit (by process name)
            WaitForMainProcessExit(mainExe, 30000);

            if (!File.Exists(zipFile))
            {
                Log($"ZIP file not found: {zipFile}");
                return 2;
            }

            try
            {
                ExtractZipWithOverwrite(zipFile, appFolder);

                try
                {
                    File.Delete(zipFile);
                }
                catch (Exception ex)
                {
                    Log("Không thể xóa file zip: " + ex.Message);
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = mainExe,
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(mainExe)
                });

                Log("Updater finished successfully.");
                return 0;
            }
            catch (Exception ex)
            {
                Log("Update lỗi: " + ex.ToString());
                return 3;
            }
        }
        catch (Exception ex)
        {
            Log("Unhandled exception: " + ex.ToString());
            return 100;
        }
    }
    static void WaitForMainProcessExit(string mainExePath, int timeoutMs)
    {
        try
        {
            string processName = Path.GetFileNameWithoutExtension(mainExePath);
            if (string.IsNullOrEmpty(processName))
            {
                Log("Không thể lấy process name từ mainExePath: " + mainExePath);
                return;
            }

            var processes = Process.GetProcessesByName(processName);
            if (processes.Length == 0)
            {
                Log($"Không tìm thấy tiến trình '{processName}' đang chạy.");
                return;
            }

            Log($"Đang chờ {processes.Length} tiến trình '{processName}' thoát (timeout {timeoutMs}ms).");
            foreach (var proc in processes)
            {
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.WaitForExit(timeoutMs);
                        Log($"Process Id {proc.Id} exited: {proc.HasExited}");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Lỗi khi chờ process {proc.Id}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Log("WaitForMainProcessExit lỗi: " + ex.Message);
        }
    }

    static void ExtractZipWithOverwrite(string zipPath, string extractPath)
    {
        using (ZipArchive archive = ZipFile.OpenRead(zipPath))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                string destinationPath = Path.Combine(extractPath, entry.FullName);
                try
                {
                    string dir = Path.GetDirectoryName(destinationPath);
                    if (!string.IsNullOrEmpty(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    if (string.IsNullOrEmpty(entry.Name))
                    {
                        // directory entry
                        continue;
                    }

                    if (File.Exists(destinationPath))
                    {
                        try { File.SetAttributes(destinationPath, FileAttributes.Normal); } catch { }
                        try { File.Delete(destinationPath); } catch (Exception ex) { Log($"Không thể xóa {destinationPath}: {ex.Message}"); }
                    }

                    entry.ExtractToFile(destinationPath, overwrite: true);
                    Log($"Extracted: {entry.FullName} -> {destinationPath}");
                }
                catch (Exception ex)
                {
                    Log($"Lỗi khi giải nén file {entry.FullName}: {ex}");
                }
            }
        }
    }

    static void Log(string message)
    {
        try
        {
            string logFile = Path.Combine(Path.GetTempPath(), "Updater.log");
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
            File.AppendAllText(logFile, line);
        }
        catch
        {
            // ignore logging errors
        }
    }
}