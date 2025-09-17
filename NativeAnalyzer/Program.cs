using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
// using System.Runtime.InteropServices; (already included below)
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Management;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Net.Http;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;

namespace NativeAnalyzer
{
    class Program
    {
        static string? currentLogFile;

        static void Log(string message)
        {
            try
            {
                var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}" + Environment.NewLine;
                Console.WriteLine(line);
                if (!string.IsNullOrEmpty(currentLogFile)) File.AppendAllText(currentLogFile, line);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    // ShowWindow declared once below; SW_RESTORE constant kept here
    private const int SW_RESTORE = 9;
    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")] 
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
            public MEMORYSTATUSEX() { dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX)); }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);

        static ulong GetTotalPhysicalMemory()
        {
            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem)) return mem.ullTotalPhys;
            return 0;
        }

        static ulong GetAvailableMemory()
        {
            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem)) return mem.ullAvailPhys;
            return 0;
        }

        static async Task<int> Main(string[] args)
        {
            var baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Ace Network Result");
            var dateStamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var outputFolder = Path.Combine(baseFolder, dateStamp);
            Directory.CreateDirectory(outputFolder);
            currentLogFile = Path.Combine(outputFolder, "native_analyzer.log");
            Log($"Starting NativeAnalyzer at {DateTime.Now}");

            // production flow continues

            Console.WriteLine("Capturing initial system information...");
            var systemInfo = GetSystemInfo(Path.Combine(outputFolder, "SystemInfo.txt"));

            Console.WriteLine("Performing speed test...");
            var speedTest = await StartSpeedTest(Path.Combine(outputFolder, "speedtest_results.txt"));

            var targets = new[] { "RDGCHG.myrealdata.net", "RDGHTN.myrealdata.net", "RDGNV.myrealdata.net", "RDGATL.myrealdata.net" };
            var tcpingResults = new List<TcpingResult>();

            foreach (var target in targets)
            {
                var r = await StartParallelTests(target, outputFolder, 60);
                if (r != null) tcpingResults.Add(r);
            }

            Console.WriteLine("Creating summary report...");
            CreateSummaryReport(systemInfo, speedTest, tcpingResults, Path.Combine(outputFolder, "summary_report.txt"));
            Log($"Analysis complete. Results saved to: {outputFolder}");
            Console.WriteLine($"Analysis complete. Results saved to: {outputFolder}");
            return 0;
        }

    public static SystemInfo GetSystemInfo(string outputFile)
        {
            var sb = new StringBuilder();
            sb.AppendLine("*****************************************************************************");
            sb.AppendLine("* SYSTEM INFORMATION *");
            sb.AppendLine("*****************************************************************************\n");

            // OS info
            var osInfo = "";
            double cpuLoad = 0;
            TimeSpan uptime = TimeSpan.Zero;
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    osInfo = $"OS Name: {obj["Caption"]}\nVersion: {obj["Version"]}\nBuild Number: {obj["BuildNumber"]}\nArchitecture: {obj["OSArchitecture"]}";
                    var lastBoot = ManagementDateTimeConverter.ToDateTime(obj["LastBootUpTime"].ToString());
                    uptime = DateTime.Now - lastBoot;
                    sb.AppendLine("OPERATING SYSTEM");
                    sb.AppendLine("----------------");
                    sb.AppendLine(osInfo);
                    sb.AppendLine();
                    sb.AppendLine("COMPUTER UPTIME");
                    sb.AppendLine("---------------");
                    sb.AppendLine($"Days: {uptime.Days}");
                    sb.AppendLine($"Hours: {uptime.Hours}");
                    sb.AppendLine($"Minutes: {uptime.Minutes}");
                    sb.AppendLine($"Seconds: {uptime.Seconds}");
                    sb.AppendLine();
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Error retrieving OS info: " + ex.Message);
            }

            // RAM info
            double totalRam = 0, freeRam = 0, usedRam = 0, ramPct = 0;
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject obj in searcher.Get())
                {
                    totalRam = Math.Round(Convert.ToDouble(obj["TotalVisibleMemorySize"]) / 1024 / 1024, 2);
                    freeRam = Math.Round(Convert.ToDouble(obj["FreePhysicalMemory"]) / 1024 / 1024, 2);
                    usedRam = Math.Round(totalRam - freeRam, 2);
                    ramPct = totalRam > 0 ? Math.Round((usedRam / totalRam) * 100, 2) : 0;
                }
                sb.AppendLine("RAM UTILIZATION");
                sb.AppendLine("--------------");
                sb.AppendLine($"Total RAM: {totalRam} GB");
                sb.AppendLine($"Used RAM: {usedRam} GB");
                sb.AppendLine($"Free RAM: {freeRam} GB");
                sb.AppendLine($"RAM Usage: {ramPct}%");
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine("Error retrieving RAM info: " + ex.Message);
            }

            // CPU info
            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                foreach (ManagementObject obj in searcher.Get())
                {
                    cpuLoad += Convert.ToDouble(obj["LoadPercentage"]);
                }
                sb.AppendLine("CPU UTILIZATION");
                sb.AppendLine("--------------");
                sb.AppendLine($"CPU Usage: {cpuLoad}%");
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine("Error retrieving CPU info: " + ex.Message);
            }

            // Disk info
            try
            {
                sb.AppendLine("STORAGE UTILIZATION");
                sb.AppendLine("------------------");
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3");
                foreach (ManagementObject obj in searcher.Get())
                {
                    var driveLetter = obj["DeviceID"];
                    var totalSpace = Math.Round(Convert.ToDouble(obj["Size"]) / 1024 / 1024 / 1024, 2);
                    var freeSpace = Math.Round(Convert.ToDouble(obj["FreeSpace"]) / 1024 / 1024 / 1024, 2);
                    var usedSpace = Math.Round(totalSpace - freeSpace, 2);
                    var usedPct = totalSpace > 0 ? Math.Round((usedSpace / totalSpace) * 100, 2) : 0;
                    sb.AppendLine($"Drive {driveLetter}");
                    sb.AppendLine($"  Total Space: {totalSpace} GB");
                    sb.AppendLine($"  Used Space: {usedSpace} GB");
                    sb.AppendLine($"  Free Space: {freeSpace} GB");
                    sb.AppendLine($"  Usage: {usedPct}%");
                    sb.AppendLine();
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("Error retrieving disk info: " + ex.Message);
            }

            File.WriteAllText(outputFile, sb.ToString());
            return new SystemInfo { RAMUsage = ramPct, CPUUsage = cpuLoad, Uptime = uptime };
        }

        // Take screenshot on a dedicated STA thread to avoid handle errors
    public static void TakeScreenshot(string filePath)
        {
            Log($"Attempting to capture screenshot to {filePath}");
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var tcs = new TaskCompletionSource<bool>();
            var thread = new Thread(() =>
            {
                try
                {
                    Application.SetHighDpiMode(HighDpiMode.SystemAware);
                    // Create invisible form to establish UI thread context
                    using var form = new Form { ShowInTaskbar = false, FormBorderStyle = FormBorderStyle.None, Opacity = 0 };
                    form.Load += (s, e) => { };
                    form.Show();
                    Application.DoEvents();

                    // Try to focus and maximize WinMTR window if present (use ShowWindow+SetForegroundWindow)
                    var processes = Process.GetProcessesByName("WinMTR");
                    IntPtr focusedWindow = IntPtr.Zero;
                    if (processes.Length > 0)
                    {
                        var winmtr = processes[0];
                        var h = winmtr.MainWindowHandle;
                        if (h != IntPtr.Zero)
                        {
                            try
                            {
                                // Restore then maximize
                                ShowWindow(h, SW_RESTORE);
                                SetForegroundWindow(h);
                                Thread.Sleep(300);
                                ShowWindow(h, 3); // SW_MAXIMIZE = 3
                                Thread.Sleep(500);
                                focusedWindow = h;
                            }
                            catch (Exception fx)
                            {
                                Log($"Focus attempt failed: {fx.GetType().Name} - {fx.Message}");
                            }
                        }
                    }

                    var allScreens = Screen.AllScreens;
                    int totalWidth = allScreens.Sum(s => s.Bounds.Width);
                    int maxHeight = allScreens.Max(s => s.Bounds.Height);

                    using var bmp = new Bitmap(totalWidth, maxHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    using var g = Graphics.FromImage(bmp);
                    int offsetX = 0;
                    bool copySuccess = true;
                    try
                    {
                        foreach (var screen in allScreens)
                        {
                            g.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, offsetX, 0, screen.Bounds.Size);
                            offsetX += screen.Bounds.Width;
                        }
                    }
                    catch (System.ComponentModel.Win32Exception wex)
                    {
                        copySuccess = false;
                        Log($"CopyFromScreen Win32Exception: {wex.Message}");
                    }

                    // If CopyFromScreen failed, attempt PrintWindow on the focused WinMTR window
                    if (!copySuccess)
                    {
                        try
                        {
                            IntPtr h = focusedWindow;
                            if (h == IntPtr.Zero)
                            {
                                var winmtrProcesses = Process.GetProcessesByName("WinMTR");
                                if (winmtrProcesses.Length > 0) h = winmtrProcesses[0].MainWindowHandle;
                            }

                            if (h != IntPtr.Zero)
                            {
                                if (GetWindowRect(h, out var r))
                                {
                                    var w = r.Right - r.Left;
                                    var hgh = r.Bottom - r.Top;
                                    using var single = new Bitmap(w, hgh, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                                    using var g2 = Graphics.FromImage(single);
                                    var hdc = g2.GetHdc();
                                    try { PrintWindow(h, hdc, 0); }
                                    finally { g2.ReleaseHdc(hdc); }
                                    // Draw captured window into full-size bitmap at 0,0
                                    g.DrawImage(single, 0, 0, totalWidth, maxHeight);
                                    Log("Captured window with PrintWindow fallback (fullscreen)");
                                }
                            }
                        }
                        catch (Exception px)
                        {
                            Log($"PrintWindow fallback failed: {px.GetType().Name} - {px.Message}");
                        }
                    }
                    // Save and if fails due to Win32Exception, retry once after a short delay
                    try
                    {
                        bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                        Log($"Screenshot saved to {filePath}");
                    }
                    catch (System.ComponentModel.Win32Exception wex)
                    {
                        Log($"Save Win32Exception, retrying: {wex.Message}");
                        Thread.Sleep(250);
                        bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                        Log($"Screenshot saved to {filePath} (retry)");
                    }
                    form.Close();
                    tcs.SetResult(true);
                }
                catch (Exception ex)
                {
                    Log($"Screenshot error: {ex.GetType().Name} - {ex.Message}");
                    Log(ex.StackTrace ?? string.Empty);
                    tcs.SetException(ex);
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            try { tcs.Task.Wait(TimeSpan.FromSeconds(10)); } catch { }
        }
    public static async Task<SpeedTestResult?> StartSpeedTest(string outputFile)
        {
            // Try to find speedtest.exe in common locations
            var candidates = new[] {
            Path.Combine("C:", "Program Files", "Speedtest", "speedtest.exe"),
            Path.Combine("C:", "Program Files (x86)", "Speedtest", "speedtest.exe"),
            Path.Combine(AppContext.BaseDirectory, "speedtest.exe")
        };

            string? exe = null;
            foreach (var c in candidates) if (File.Exists(c)) { exe = c; break; }

            // Try to find in PATH using 'where'
            if (exe == null)
            {
                try
                {
                    var wherePsi = new ProcessStartInfo("where", "speedtest.exe") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                    using var whereProc = Process.Start(wherePsi);
                    if (whereProc != null)
                    {
                        var whereOut = await whereProc.StandardOutput.ReadToEndAsync();
                        whereProc.WaitForExit();
                        var first = whereOut.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                        if (!string.IsNullOrEmpty(first) && File.Exists(first)) exe = first.Trim();
                    }
                }
                catch { }
            }

            // If still not found, try to download the CLI (same behavior as PowerShell script)
            if (exe == null)
            {
                Console.WriteLine("Speedtest CLI not found. Attempting to download...");
                var speedtestUrl = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip";
                var tempZipFile = Path.Combine(Path.GetTempPath(), "speedtest.zip");
                var tempExtractPath = Path.Combine(Path.GetTempPath(), "speedtest");

                try
                {
                    using var http = new HttpClient();
                    using var resp = await http.GetAsync(speedtestUrl);
                    resp.EnsureSuccessStatusCode();
                    await using (var fs = File.Create(tempZipFile))
                    {
                        await resp.Content.CopyToAsync(fs);
                    }

                    if (Directory.Exists(tempExtractPath)) Directory.Delete(tempExtractPath, true);
                    Directory.CreateDirectory(tempExtractPath);
                    ZipFile.ExtractToDirectory(tempZipFile, tempExtractPath);

                    var found = Directory.GetFiles(tempExtractPath, "speedtest.exe", SearchOption.AllDirectories).FirstOrDefault();
                    if (!string.IsNullOrEmpty(found)) exe = found;
                    else Console.WriteLine("Downloaded speedtest but executable not found inside archive.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to download Speedtest CLI: " + ex.Message);
                }
            }

            if (exe == null)
            {
                Console.WriteLine("Speedtest CLI not available. Skipping speedtest.");
                return null;
            }

            try
            {
                var psi = new ProcessStartInfo(exe, "--accept-license --accept-gdpr --format=json")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var p = Process.Start(psi);
                if (p == null) return null;
                var output = await p.StandardOutput.ReadToEndAsync();
                p.WaitForExit();
                File.WriteAllText(outputFile, output);

                try
                {
                    using var json = JsonDocument.Parse(output);
                    var root = json.RootElement;

                    double downloadBandwidth = 0, uploadBandwidth = 0, ping = 0, jitter = 0, packetLoss = 0;

                    if (root.TryGetProperty("download", out var dl) && dl.ValueKind == JsonValueKind.Object)
                    {
                        if (dl.TryGetProperty("bandwidth", out var bw) && bw.TryGetDouble(out var bwd)) downloadBandwidth = bwd;
                    }
                    if (root.TryGetProperty("upload", out var ul) && ul.ValueKind == JsonValueKind.Object)
                    {
                        if (ul.TryGetProperty("bandwidth", out var bw) && bw.TryGetDouble(out var bwu)) uploadBandwidth = bwu;
                    }
                    if (root.TryGetProperty("ping", out var pingEl) && pingEl.ValueKind == JsonValueKind.Object)
                    {
                        if (pingEl.TryGetProperty("latency", out var pl) && pl.TryGetDouble(out var pd)) ping = pd;
                        if (pingEl.TryGetProperty("jitter", out var pj) && pj.TryGetDouble(out var jd)) jitter = jd;
                    }
                    if (root.TryGetProperty("packetLoss", out var ploss) && ploss.TryGetDouble(out var pld)) packetLoss = pld;

                    var download = downloadBandwidth * 8 / 1000000.0;
                    var upload = uploadBandwidth * 8 / 1000000.0;

                    return new SpeedTestResult { Download = Math.Round(download, 2), Upload = Math.Round(upload, 2), Ping = ping, Jitter = jitter, PacketLoss = packetLoss };
                }
                catch (Exception ex)
                {
                    File.AppendAllText(outputFile, "\n// Error parsing JSON: " + ex.ToString());
                    Console.WriteLine("Error parsing speedtest JSON: " + ex.Message);
                    return null;
                }
            }
            catch (Exception ex)
            {
                File.WriteAllText(outputFile, "Error: " + ex.ToString());
                Console.WriteLine("Error running speedtest: " + ex.Message);
                return null;
            }
        }

    public static async Task<TcpingResult?> StartParallelTests(string target, string outputFolder, int testDurationSeconds)
    {
        var tcpingPath = Path.Combine(AppContext.BaseDirectory, "tcping.exe");
        var winmtrPath = Path.Combine(AppContext.BaseDirectory, "WinMTR.exe");

        var tcpingOutput = Path.Combine(outputFolder, target + "_tcping.txt");

            if (!File.Exists(tcpingPath))
        {
            // try to find in PATH
            try
            {
                var wherePsi = new ProcessStartInfo("where", "tcping.exe") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                using var whereProc = Process.Start(wherePsi);
                if (whereProc != null)
                {
                    var whereOut = await whereProc.StandardOutput.ReadToEndAsync();
                    whereProc.WaitForExit();
                    var first = whereOut.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrEmpty(first) && File.Exists(first)) tcpingPath = first.Trim();
                }
            }
            catch { }
        }

        if (!File.Exists(tcpingPath))
        {
            Console.WriteLine("tcping.exe not found; TCPing will be skipped for " + target);
            File.WriteAllText(tcpingOutput, "TCPing executable not found. Test skipped.");
        }
        else
        {
            var tcpProc = new ProcessStartInfo(tcpingPath, $"-n 30 {target} 443")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var p = Process.Start(tcpProc);
            if (p != null)
            {
                var outStr = await p.StandardOutput.ReadToEndAsync();
                File.WriteAllText(tcpingOutput, outStr);
                p.WaitForExit();
            }
        }

        if (File.Exists(winmtrPath))
        {
            Log($"Starting WinMTR from {winmtrPath}");
            // Start WinMTR normally so its window is visible for screenshots
            var winmtrProc = new ProcessStartInfo(winmtrPath, target)
            {
                UseShellExecute = true,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal
            };

            var proc = Process.Start(winmtrProc);
            if (proc == null)
            {
                Log("Failed to start WinMTR process");
                return null;
            }
            // Do not hide the WinMTR window — allow it to be visible so screenshots work
            
            Log($"Running WinMTR for {target} for {testDurationSeconds} seconds...");
            
            // Give WinMTR time to start collecting data
            await Task.Delay((testDurationSeconds - 5) * 1000);
            
            // Take screenshot when data is collected
            var screenshotPath = Path.Combine(outputFolder, target + "_winmtr.png");
            Log($"Attempting to capture screenshot to {screenshotPath}");
            TakeScreenshot(screenshotPath);
            
            // Wait remaining time and cleanup
            await Task.Delay(5000);
            try 
            { 
                if (!proc.HasExited)
                {
                    Log("Terminating WinMTR process");
                    proc.Kill(); 
                }
            } 
            catch (Exception ex) 
            { 
                Log($"Error terminating WinMTR: {ex.Message}");
            }
        }
        else
        {
            Log($"WinMTR executable not found at {winmtrPath}");
        }

        string tcpData = File.Exists(tcpingOutput) ? File.ReadAllText(tcpingOutput) : string.Empty;
        return new TcpingResult { Target = target, TCPingData = tcpData };
    }

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_HIDE = 0;

    private static void TryHideProcessWindow(Process proc, int timeoutMilliseconds = 2000)
    {
        try
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMilliseconds)
            {
                proc.Refresh();
                var h = proc.MainWindowHandle;
                if (h != IntPtr.Zero)
                {
                    ShowWindow(h, SW_HIDE);
                    break;
                }
                Thread.Sleep(100);
            }
        }
        catch
        {
            // best-effort; ignore any failures
        }
    }

    

    public static void CreateSummaryReport(SystemInfo sys, SpeedTestResult? speed, List<TcpingResult> tcpResults, string outputFile)
    {
        var sb = new StringBuilder();
        sb.AppendLine(new string('*', 77));
        sb.AppendLine("* SUMMARY REPORT *");
        sb.AppendLine(new string('*', 77));
        sb.AppendLine();

        // SYSTEM HEALTH
        sb.AppendLine("SYSTEM HEALTH CHECK");
        sb.AppendLine("------------------");
        if (sys.RAMUsage > 80 || sys.CPUUsage > 80)
        {
            if (sys.RAMUsage > 80) sb.AppendLine($"! WARNING: High RAM usage ({sys.RAMUsage}%)");
            if (sys.CPUUsage > 80) sb.AppendLine($"! WARNING: High CPU usage ({sys.CPUUsage}%)");
        }
        else
        {
            sb.AppendLine("✓ No system issues detected.");
        }
        sb.AppendLine();

        // INTERNET
        sb.AppendLine("INTERNET CONNECTION CHECK");
        sb.AppendLine("------------------------");
        if (speed == null)
        {
            sb.AppendLine("ERROR: Could not perform speed test.");
        }
        else
        {
            sb.AppendLine("✓ Internet connection appears healthy.");
            sb.AppendLine($"  Download: {speed.Download} Mbps");
            sb.AppendLine($"  Upload: {speed.Upload} Mbps");
            sb.AppendLine($"  Ping: {Math.Round(speed.Ping, 3)} ms");
        }
        sb.AppendLine();

        // SERVER CONNECTIVITY
        sb.AppendLine("SERVER CONNECTIVITY CHECK");
        sb.AppendLine("-----------------------");

        if (tcpResults == null || tcpResults.Count == 0)
        {
            sb.AppendLine("! No TCPing results available.");
            File.WriteAllText(outputFile, sb.ToString());
            return;
        }

        // Parse ping times for each server
        var serverStats = new List<(string Target, double? Avg, double? Min, double? Max)>();
        foreach (var r in tcpResults)
        {
            if (string.IsNullOrWhiteSpace(r.TCPingData))
            {
                serverStats.Add((r.Target, null, null, null));
                continue;
            }

            var times = ParsePingTimes(r.TCPingData);
            if (times.Count == 0) serverStats.Add((r.Target, null, null, null));
            else serverStats.Add((r.Target, times.Average(), times.Min(), times.Max()));
        }

        // Best performing server = lowest average
        var best = serverStats.Where(s => s.Avg.HasValue).OrderBy(s => s.Avg.Value).FirstOrDefault();
        if (best.Target != null && best.Avg.HasValue)
        {
            sb.AppendLine($"BEST PERFORMING SERVER: {best.Target}");
            sb.AppendLine($"  Average ping time: {Math.Round(best.Avg.Value, 2)} ms");
            sb.AppendLine($"  Min/Max ping time: {Math.Round(best.Min.Value, 2)} / {Math.Round(best.Max.Value, 2)} ms");
        }
        else
        {
            sb.AppendLine("BEST PERFORMING SERVER: N/A");
        }

        sb.AppendLine();
        sb.AppendLine("ALL SERVER RESULTS:");
        foreach (var s in serverStats)
        {
            if (!s.Avg.HasValue)
            {
                sb.AppendLine($"- {s.Target}: No data");
            }
            else
            {
                sb.AppendLine($"- {s.Target}: Avg={Math.Round(s.Avg.Value, 2)} ms, Min={Math.Round(s.Min.Value, 2)} ms, Max={Math.Round(s.Max.Value, 2)} ms");
            }
        }

        File.WriteAllText(outputFile, sb.ToString());
    }

    // Parse potential ping times from tcping or other output. Looks for numbers followed by 'ms' or standalone numbers.
    static List<double> ParsePingTimes(string text)
    {
        var results = new List<double>();
        if (string.IsNullOrWhiteSpace(text)) return results;

        try
        {
            // Find numbers followed by 'ms'
            var msRegex = new System.Text.RegularExpressions.Regex(@"(\d+\.?\d*)\s*ms", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            foreach (System.Text.RegularExpressions.Match m in msRegex.Matches(text))
            {
                if (double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) results.Add(v);
            }

            if (results.Count > 0) return results;

            // Fallback: find standalone numbers that look like ms values (small numbers)
            var numRegex = new System.Text.RegularExpressions.Regex(@"(\d+\.?\d*)");
            foreach (System.Text.RegularExpressions.Match m in numRegex.Matches(text))
            {
                if (double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                {
                    if (v >= 0 && v < 10000) results.Add(v);
                }
            }
        }
        catch { }

        return results;
    }

    // Duplicate record types removed. Use the public records declared in SystemInfo.cs
}
}
