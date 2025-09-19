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
using System.Drawing.Imaging;

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

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    static void BringCurrentProcessWindowToFront()
    {
        try
        {
            var currentPid = (uint)Process.GetCurrentProcess().Id;
            IntPtr found = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true; // continue enumerating
                GetWindowThreadProcessId(hWnd, out var pid);
                if (pid == currentPid)
                {
                    found = hWnd;
                    return false; // stop enumeration
                }
                return true;
            }, IntPtr.Zero);

            if (found != IntPtr.Zero)
            {
                // Restore and bring to front. Toggle topmost briefly to ensure visibility over new windows.
                ShowWindow(found, SW_RESTORE);
                // Make topmost
                try { SetWindowPos(found, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW); } catch { }
                SetForegroundWindow(found);
                Thread.Sleep(300);
                // Remove topmost so window behaves normally
                try { SetWindowPos(found, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW); } catch { }
                Thread.Sleep(150);
            }
        }
        catch { }
    }

    // ShowWindow declared once below; SW_RESTORE constant kept here
    private const int SW_RESTORE = 9;
    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    private static readonly IntPtr HWND_TOP = IntPtr.Zero;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;

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

            var targets = new[] { "RDGCHG.myrealdata.net", "RDGHTN.myrealdata.net", "RDGNV.myrealdata.net", "RDGATL.myrealdata.net", "RDGDEN.myrealdata.net" };

            // Announce start of gateway scanning
            var startMsg = "Starting Ace Cloud Gateway Servers Scanning";
            Console.WriteLine(startMsg);
            Log(startMsg);

            // Run tcping and WinMTR tests concurrently across all targets
            var tcpingResults = await RunConcurrentGatewayScans(targets.ToList(), outputFolder, 60);

            // Announce completion of gateway scanning and move to report generation
            var genMsg = "Generating Analysis Reports";
            Console.WriteLine(genMsg);
            Log(genMsg);

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

                // Previously we saved raw JSON; now we only produce the human-readable text report.

                try
                {
                    using var json = JsonDocument.Parse(output);
                    var root = json.RootElement;

                    double downloadBandwidth = 0, uploadBandwidth = 0, ping = 0, jitter = 0, packetLoss = 0;
                    string isp = "";
                    string internalIp = "";
                    string externalIp = "";
                    string interfaceName = "";
                    string macAddr = "";
                    bool isVpn = false;
                    string serverHost = "";
                    string serverName = "";
                    string serverLocation = "";
                    string serverCountry = "";
                    string serverIp = "";
                    int serverPort = 0;
                    string serverId = "";
                    string resultUrl = "";

                    // latency details
                    double dl_iqm = 0, dl_low = 0, dl_high = 0, dl_jitter = 0;
                    double ul_iqm = 0, ul_low = 0, ul_high = 0, ul_jitter = 0;

                    if (root.TryGetProperty("download", out var dl) && dl.ValueKind == JsonValueKind.Object)
                    {
                        if (dl.TryGetProperty("bandwidth", out var bw) && bw.TryGetDouble(out var bwd)) downloadBandwidth = bwd;
                        if (dl.TryGetProperty("latency", out var dlat) && dlat.ValueKind == JsonValueKind.Object)
                        {
                            if (dlat.TryGetProperty("iqm", out var iq) && iq.TryGetDouble(out var iqv)) dl_iqm = iqv;
                            if (dlat.TryGetProperty("low", out var low) && low.TryGetDouble(out var lowv)) dl_low = lowv;
                            if (dlat.TryGetProperty("high", out var high) && high.TryGetDouble(out var highv)) dl_high = highv;
                            if (dlat.TryGetProperty("jitter", out var dj) && dj.TryGetDouble(out var djv)) dl_jitter = djv;
                        }
                    }
                    if (root.TryGetProperty("upload", out var ul) && ul.ValueKind == JsonValueKind.Object)
                    {
                        if (ul.TryGetProperty("bandwidth", out var bw) && bw.TryGetDouble(out var bwu)) uploadBandwidth = bwu;
                        if (ul.TryGetProperty("latency", out var ulat) && ulat.ValueKind == JsonValueKind.Object)
                        {
                            if (ulat.TryGetProperty("iqm", out var iq) && iq.TryGetDouble(out var iqv)) ul_iqm = iqv;
                            if (ulat.TryGetProperty("low", out var low) && low.TryGetDouble(out var lowv)) ul_low = lowv;
                            if (ulat.TryGetProperty("high", out var high) && high.TryGetDouble(out var highv)) ul_high = highv;
                            if (ulat.TryGetProperty("jitter", out var uj) && uj.TryGetDouble(out var ujv)) ul_jitter = ujv;
                        }
                    }
                    if (root.TryGetProperty("ping", out var pingEl) && pingEl.ValueKind == JsonValueKind.Object)
                    {
                        if (pingEl.TryGetProperty("latency", out var pl) && pl.TryGetDouble(out var pd)) ping = pd;
                        if (pingEl.TryGetProperty("jitter", out var pj) && pj.TryGetDouble(out var jd)) jitter = jd;
                    }
                    if (root.TryGetProperty("packetLoss", out var ploss) && ploss.TryGetDouble(out var pld)) packetLoss = pld;
                    if (root.TryGetProperty("isp", out var ispEl) && ispEl.ValueKind == JsonValueKind.String) isp = ispEl.GetString() ?? "";
                    if (root.TryGetProperty("interface", out var ifEl) && ifEl.ValueKind == JsonValueKind.Object)
                    {
                        if (ifEl.TryGetProperty("internalIp", out var iip) && iip.ValueKind == JsonValueKind.String) internalIp = iip.GetString() ?? "";
                        if (ifEl.TryGetProperty("externalIp", out var eip) && eip.ValueKind == JsonValueKind.String) externalIp = eip.GetString() ?? "";
                        if (ifEl.TryGetProperty("name", out var iname) && iname.ValueKind == JsonValueKind.String) interfaceName = iname.GetString() ?? "";
                        if (ifEl.TryGetProperty("macAddr", out var mac) && mac.ValueKind == JsonValueKind.String) macAddr = mac.GetString() ?? "";
                        if (ifEl.TryGetProperty("isVpn", out var vpn) && vpn.ValueKind == JsonValueKind.True) isVpn = true;
                    }
                    if (root.TryGetProperty("server", out var srv) && srv.ValueKind == JsonValueKind.Object)
                    {
                        if (srv.TryGetProperty("host", out var sh) && sh.ValueKind == JsonValueKind.String) serverHost = sh.GetString() ?? "";
                        if (srv.TryGetProperty("name", out var sn) && sn.ValueKind == JsonValueKind.String) serverName = sn.GetString() ?? "";
                        if (srv.TryGetProperty("location", out var sl) && sl.ValueKind == JsonValueKind.String) serverLocation = sl.GetString() ?? "";
                        if (srv.TryGetProperty("country", out var sc) && sc.ValueKind == JsonValueKind.String) serverCountry = sc.GetString() ?? "";
                        if (srv.TryGetProperty("ip", out var sip) && sip.ValueKind == JsonValueKind.String) serverIp = sip.GetString() ?? "";
                        if (srv.TryGetProperty("port", out var sport) && sport.TryGetInt32(out var sp)) serverPort = sp;
                        if (srv.TryGetProperty("id", out var sid) && sid.ValueKind == JsonValueKind.Number) serverId = sid.ToString();
                    }
                    if (root.TryGetProperty("result", out var res) && res.ValueKind == JsonValueKind.Object)
                    {
                        if (res.TryGetProperty("url", out var ru) && ru.ValueKind == JsonValueKind.String) resultUrl = ru.GetString() ?? "";
                        // intentionally not saving result id into the text report per request
                    }

                    var download = downloadBandwidth * 8 / 1000000.0;
                    var upload = uploadBandwidth * 8 / 1000000.0;

                    // Build human-readable report
                    var sb = new StringBuilder();
                    sb.AppendLine("*****************************************************************************");
                    sb.AppendLine("* SPEEDTEST RESULTS *");
                    sb.AppendLine("*****************************************************************************\n");
                    sb.AppendLine($"TEST TIME: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine();
                    sb.AppendLine("SERVER INFORMATION");
                    sb.AppendLine("------------------");
                    // Only include non-sensitive server information: name and location
                    if (!string.IsNullOrEmpty(serverName)) sb.AppendLine($"Name: {serverName}");
                    if (!string.IsNullOrEmpty(serverLocation) || !string.IsNullOrEmpty(serverCountry)) sb.AppendLine($"Location: {serverLocation}{(string.IsNullOrEmpty(serverLocation) ? "" : ", ")}{serverCountry}");
                    sb.AppendLine();
                    sb.AppendLine("CONNECTION INFORMATION");
                    sb.AppendLine("----------------------");
                    // Only include useful connection info; avoid exposing MAC, VPN, or IP addresses
                    if (!string.IsNullOrEmpty(isp)) sb.AppendLine($"ISP: {isp}");
                    sb.AppendLine();
                    sb.AppendLine("SPEED RESULTS");
                    sb.AppendLine("-------------");
                    sb.AppendLine($"Download : {Math.Round(download, 2)} Mbps");
                    sb.AppendLine($"Upload   : {Math.Round(upload, 2)} Mbps");
                    sb.AppendLine($"Ping     : {Math.Round(ping, 2)} ms");
                    sb.AppendLine($"Jitter   : {Math.Round(jitter, 2)} ms");
                    sb.AppendLine($"Packet Loss: {Math.Round(packetLoss, 2)} %");
                    sb.AppendLine();
                    sb.AppendLine("DOWNLOAD LATENCY (ms)");
                    sb.AppendLine($"  IQM : {Math.Round(dl_iqm, 3)}");
                    sb.AppendLine($"  Low : {Math.Round(dl_low, 3)}");
                    sb.AppendLine($"  High: {Math.Round(dl_high, 3)}");
                    sb.AppendLine($"  Jitter: {Math.Round(dl_jitter, 3)}");
                    sb.AppendLine();
                    sb.AppendLine("UPLOAD LATENCY (ms)");
                    sb.AppendLine($"  IQM : {Math.Round(ul_iqm, 3)}");
                    sb.AppendLine($"  Low : {Math.Round(ul_low, 3)}");
                    sb.AppendLine($"  High: {Math.Round(ul_high, 3)}");
                    sb.AppendLine($"  Jitter: {Math.Round(ul_jitter, 3)}");
                    sb.AppendLine();
                    if (!string.IsNullOrEmpty(resultUrl)) sb.AppendLine($"Result URL: {resultUrl}");

                    File.WriteAllText(outputFile, sb.ToString());

                    return new SpeedTestResult { Download = Math.Round(download, 2), Upload = Math.Round(upload, 2), Ping = Math.Round(ping, 2), Jitter = Math.Round(jitter, 2), PacketLoss = Math.Round(packetLoss, 2) };
                }
                catch (Exception ex)
                {
                    // If parsing fails, preserve raw output to the requested file so user still sees data
                    try { File.WriteAllText(outputFile, output); } catch { }
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

    public static async Task<List<TcpingResult>> RunConcurrentGatewayScans(List<string> gateways, string outputFolder, int testDurationSeconds, IProgress<int>? percentProgress = null, IProgress<string>? statusProgress = null)
    {
        var results = new List<TcpingResult>();

        var tcpingPath = Path.Combine(AppContext.BaseDirectory, "tcping.exe");
        var winmtrPath = Path.Combine(AppContext.BaseDirectory, "WinMTR.exe");

        async Task<string?> FindExe(string defaultPath, string exeName)
        {
            if (File.Exists(defaultPath)) return defaultPath;
            try
            {
                var wherePsi = new ProcessStartInfo("where", exeName) { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                using var whereProc = Process.Start(wherePsi);
                if (whereProc != null)
                {
                    var outStr = await whereProc.StandardOutput.ReadToEndAsync();
                    whereProc.WaitForExit();
                    var first = outStr.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrEmpty(first) && File.Exists(first)) return first.Trim();
                }
            }
            catch { }
            return null;
        }

        var tcpingResolved = await FindExe(tcpingPath, "tcping.exe") ?? string.Empty;
        var winmtrResolved = await FindExe(winmtrPath, "WinMTR.exe") ?? string.Empty;

        var winProcs = new Dictionary<string, Process>();

        // Start WinMTR processes first (so windows are present when tcping completes)
        if (!string.IsNullOrEmpty(winmtrResolved))
        {
            // Ensure this application window is foreground before launching WinMTR so WinMTR windows
            // don't end up above our app unexpectedly. Make app topmost while starting/tile WinMTR.
            try { SetCurrentAppTopmost(true); } catch { }

            var handleWaitTasks = new List<Task>();
            foreach (var gateway in gateways)
            {
                try
                {
                    var psi = new ProcessStartInfo(winmtrResolved, gateway)
                    {
                        UseShellExecute = true,
                        CreateNoWindow = false,
                        WindowStyle = ProcessWindowStyle.Normal
                    };
                    var proc = Process.Start(psi);
                    if (proc != null)
                    {
                        Log($"WinMTR started for {gateway} (PID: {proc.Id})");
                        // store process reference immediately
                        winProcs[gateway] = proc;

                        // Start a background task to wait for the main window handle without blocking the loop
                        handleWaitTasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                var sw = Stopwatch.StartNew();
                                while (sw.ElapsedMilliseconds < 7000)
                                {
                                    proc.Refresh();
                                    if (proc.MainWindowHandle != IntPtr.Zero) break;
                                    await Task.Delay(200);
                                }
                            }
                            catch (Exception ex) { Log($"WinMTR handle wait error for {gateway}: {ex.Message}"); }
                        }));
                    }
                    else Log($"Failed to start WinMTR for {gateway}");
                }
                catch (Exception ex)
                {
                    Log($"Exception starting WinMTR for {gateway}: {ex.Message}");
                }
            }

            // Wait for all handle-wait tasks to complete in parallel (non-blocking start above)
            try { await Task.WhenAll(handleWaitTasks); } catch { }
        }
        else
        {
            Log("WinMTR not found; skipping WinMTR launches.");
        }

        // Start tcping tasks concurrently but do not rely on their completion to take screenshots.
        // We'll run tcping with the duration specified by testDurationSeconds and capture screenshots simultaneously at the end.
        var tcpingTasks = new List<Task<(string Target, string Output)>>();
        foreach (var gateway in gateways)
        {
            if (!string.IsNullOrEmpty(tcpingResolved))
            {
                var outPath = Path.Combine(outputFolder, SanitizeFileName(gateway) + "_tcping.txt");
                tcpingTasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        // Start tcping with exact packet count equal to testDurationSeconds
                        var psi = new ProcessStartInfo(tcpingResolved, $"-n {testDurationSeconds} {gateway} 443")
                        {
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        using var p = Process.Start(psi);
                        if (p != null)
                        {
                            var o = await p.StandardOutput.ReadToEndAsync();
                            // do not wait here for exit; we'll wait for tcping tasks below
                            File.WriteAllText(outPath, o);
                            Log($"tcping started for {gateway}, output will be saved to {outPath}");
                            return (gateway, o);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"tcping error for {gateway}: {ex.Message}");
                        File.WriteAllText(outPath, "tcping error: " + ex.ToString());
                    }
                    return (gateway, string.Empty);
                }));
            }
            else
            {
                var outPath = Path.Combine(outputFolder, SanitizeFileName(gateway) + "_tcping.txt");
                File.WriteAllText(outPath, "tcping not found; skipped.");
                tcpingTasks.Add(Task.FromResult((gateway, File.ReadAllText(outPath))));
            }
        }

        // Arrange/tile WinMTR windows across the primary screen evenly
        try
        {
            var screen = Screen.PrimaryScreen.Bounds;
            int n = winProcs.Count;
            if (n > 0)
            {
                int cols = (int)Math.Ceiling(Math.Sqrt(n));
                int rows = (int)Math.Ceiling((double)n / cols);
                int w = screen.Width / cols;
                int h = screen.Height / rows;
                int idx = 0;
                foreach (var kv in winProcs)
                {
                    try
                    {
                        var p = kv.Value;
                        p.Refresh();
                        var hWnd = p.MainWindowHandle;
                        if (hWnd == IntPtr.Zero) continue;
                        int col = idx % cols;
                        int row = idx / cols;
                        int x = screen.Left + col * w;
                        int y = screen.Top + row * h;
                        SetWindowPos(hWnd, HWND_TOP, x, y, w, h, SWP_NOZORDER | SWP_SHOWWINDOW);
                        idx++;
                    }
                    catch (Exception ex) { Log($"Tiling error: {ex.Message}"); }
                }
            }
        }
        catch (Exception ex) { Log($"Tiling windows failed: {ex.Message}"); }

        // Wait for the full test duration (this will be the moment we take screenshots).
        Log($"Waiting {testDurationSeconds} seconds for tcping/WinMTR to run before taking screenshots...");
        for (int elapsed = 0; elapsed < testDurationSeconds; elapsed++)
        {
            // Report progress as percentage complete (elapsed seconds)
            try { percentProgress?.Report((int)Math.Round((double)elapsed / testDurationSeconds * 100)); } catch { }
            await Task.Delay(1000);
        }
        // final report: 100%
        try { percentProgress?.Report(100); } catch { }

        // At the end of the interval, capture all WinMTR windows simultaneously
        var captureTasks = new List<Task>();
        foreach (var kv in winProcs)
        {
            var gateway = kv.Key;
            var proc = kv.Value;
                captureTasks.Add(Task.Run(() =>
            {
                try
                {
                    proc.Refresh();
                    var h = proc.MainWindowHandle;
                    if (h == IntPtr.Zero) return;
                        // Maximize and bring to foreground so we capture full window
                        try { MaximizeWindowAndWait(h); } catch { }
                        var screenshotPath = Path.Combine(outputFolder, SanitizeFileName(gateway) + "_winmtr.png");
                        var captured = CaptureWindow(h, screenshotPath);
                    if (!captured)
                    {
                        Log($"CaptureWindow failed for {gateway}, falling back to TakeScreenshot");
                        TakeScreenshot(screenshotPath);
                    }
                }
                catch (Exception ex) { Log($"Capture error for {gateway}: {ex.Message}"); }
            }));
        }

    try { await Task.WhenAll(captureTasks); } catch { }

    // Inform UI/reporting that captures are done and we're generating reports
    try { statusProgress?.Report("Generating Analysis Reports"); } catch { }

    // Restore app topmost state to normal so user windows behave normally
    try { SetCurrentAppTopmost(false); } catch { }

        // Ensure any remaining WinMTR processes are cleaned up
        foreach (var kv in winProcs)
        {
            try { if (!kv.Value.HasExited) kv.Value.Kill(); }
            catch { }
        }

        // Wait for tcping tasks to finish and collect results
        var tcpResults = await Task.WhenAll(tcpingTasks);
        foreach (var t in tcpResults)
        {
            results.Add(new TcpingResult { Target = t.Target, TCPingData = t.Output });
        }

        return results;
    }

    // Replace CaptureWindow stub with a real implementation that saves the window specified by handle
static bool CaptureWindow(IntPtr hWnd, string filePath)
{
    try
    {
        if (hWnd == IntPtr.Zero)
        {
            Log("CaptureWindow: hWnd is zero");
            return false;
        }

        if (!GetWindowRect(hWnd, out var rect))
        {
            Log("CaptureWindow: GetWindowRect failed");
            return false;
        }

        int width = Math.Max(1, rect.Right - rect.Left);
        int height = Math.Max(1, rect.Bottom - rect.Top);

        using var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);

        var hdc = g.GetHdc();
        try
        {
            var ok = PrintWindow(hWnd, hdc, 0);
            g.ReleaseHdc(hdc);
            if (!ok)
            {
                // fallback to screen capture of the window rectangle
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
            }
        }
        catch (Exception ex)
        {
            try { g.ReleaseHdc(hdc); } catch { }
            Log($"CaptureWindow PrintWindow exception: {ex.Message}");
            // fallback to CopyFromScreen
            try { g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy); }
            catch (Exception ex2) { Log($"CaptureWindow CopyFromScreen failed: {ex2.Message}"); }
        }

        // Ensure output directory exists
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        bmp.Save(filePath, ImageFormat.Png);
        Log($"CaptureWindow saved: {filePath}");
        return true;
    }
    catch (Exception ex)
    {
        Log($"CaptureWindow failed: {ex.Message}");
        return false;
    }
}

// Cached handle to the app window used when toggling topmost state
static IntPtr cachedAppWindow = IntPtr.Zero;

static void SetCurrentAppTopmost(bool makeTopmost)
{
    try
    {
        if (cachedAppWindow == IntPtr.Zero)
        {
            var currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                GetWindowThreadProcessId(hWnd, out var pid);
                if (pid == currentPid)
                {
                    cachedAppWindow = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
        }

        if (cachedAppWindow != IntPtr.Zero)
        {
            if (makeTopmost)
            {
                SetWindowPos(cachedAppWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
                ShowWindow(cachedAppWindow, SW_RESTORE);
                SetForegroundWindow(cachedAppWindow);
            }
            else
            {
                SetWindowPos(cachedAppWindow, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            }
        }
    }
    catch { }
}

// Attempt to maximize a window and wait a short time for it to repaint/resize
static void MaximizeWindowAndWait(IntPtr hWnd)
{
    try
    {
        // SW_MAXIMIZE = 3
        ShowWindow(hWnd, 3);
        SetForegroundWindow(hWnd);
        // Give the window some time to resize and redraw
        Thread.Sleep(450);
        // Try again to ensure it's maximized
        ShowWindow(hWnd, 3);
        Thread.Sleep(200);
    }
    catch { }
}

// Utility to sanitize filenames
static string SanitizeFileName(string name)
{
    var invalid = Path.GetInvalidFileNameChars();
    var sb = new StringBuilder(name.Length);
    foreach (var c in name)
    {
        sb.Append(invalid.Contains(c) ? '_' : c);
    }
    return sb.ToString();
}

    public static void CreateSummaryReport(SystemInfo sys, SpeedTestResult? speed, List<TcpingResult> tcpResults, string outputFile)
    {
        var sb = new StringBuilder();
        sb.AppendLine(new string('*', 77));
        sb.AppendLine("* SUMMARY REPORT *");
        sb.AppendLine(new string('*', 77));
        sb.AppendLine();

        // Thresholds (user-specified)
        const double DOWNLOAD_THRESHOLD = 20.0; // Mbps
        const double UPLOAD_THRESHOLD = 20.0; // Mbps
        const double DL_LATENCY_THRESHOLD = 25.0; // ms
        const double UL_LATENCY_THRESHOLD = 25.0; // ms
        const double PING_THRESHOLD = 10.0; // ms
        const double RAM_THRESHOLD = 90.0; // percent
        const double CPU_THRESHOLD = 90.0; // percent
        const double DISK_FREE_THRESHOLD = 20.0; // percent free
        const double UPTIME_DAYS_THRESHOLD = 2.0; // days

        // SYSTEM HEALTH
    sb.AppendLine("SYSTEM HEALTH CHECK");
    sb.AppendLine(new string('=', 28));
        var systemIssues = new List<string>();
        if (sys.RAMUsage >= RAM_THRESHOLD) systemIssues.Add($"WARNING: High RAM usage ({sys.RAMUsage}%) (threshold {RAM_THRESHOLD}%)");
        if (sys.CPUUsage >= CPU_THRESHOLD) systemIssues.Add($"WARNING: High CPU usage ({sys.CPUUsage}%) (threshold {CPU_THRESHOLD}%)");
        if (sys.Uptime.TotalDays >= UPTIME_DAYS_THRESHOLD) systemIssues.Add($"WARNING: System uptime is {sys.Uptime.Days} days {sys.Uptime.Hours} hours (recommended < {UPTIME_DAYS_THRESHOLD} days)");

        // Check disk free percentage across drives
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3");
            foreach (ManagementObject obj in searcher.Get())
            {
                var drive = obj["DeviceID"]?.ToString() ?? "?";
                var sizeObj = obj["Size"];
                var freeObj = obj["FreeSpace"];
                if (sizeObj != null && freeObj != null)
                {
                    var total = Convert.ToDouble(sizeObj);
                    var free = Convert.ToDouble(freeObj);
                    var freePct = total > 0 ? Math.Round((free / total) * 100.0, 2) : 0.0;
                    if (freePct < DISK_FREE_THRESHOLD) systemIssues.Add($"WARNING: Low disk free space on {drive} ({freePct}% free, threshold {DISK_FREE_THRESHOLD}% )");
                }
            }
        }
        catch { }

        if (systemIssues.Count == 0)
        {
            sb.AppendLine("✓ No system issues detected.");
        }
        else
        {
            foreach (var it in systemIssues) sb.AppendLine("! " + it);
        }
        sb.AppendLine();

        // INTERNET
    sb.AppendLine("INTERNET CONNECTION CHECK");
    sb.AppendLine(new string('=', 24));
        if (speed == null)
        {
            sb.AppendLine("ERROR: Could not perform speed test.");
            sb.AppendLine();
        }
        else
        {
            sb.AppendLine($"  Download: {speed.Download} Mbps");
            sb.AppendLine($"  Upload: {speed.Upload} Mbps");
            sb.AppendLine($"  Ping: {Math.Round(speed.Ping, 3)} ms");
            sb.AppendLine($"  Jitter: {Math.Round(speed.Jitter, 3)} ms");
            sb.AppendLine($"  Packet Loss: {Math.Round(speed.PacketLoss, 3)} %");
            sb.AppendLine();

            // Parse speedtest_results.txt for detailed latency numbers (IQM) if present
            double? dlIQM = null, ulIQM = null;
            try
            {
                var speedTxt = Path.Combine(Path.GetDirectoryName(outputFile) ?? string.Empty, "speedtest_results.txt");
                if (File.Exists(speedTxt))
                {
                    var lines = File.ReadAllLines(speedTxt);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        var line = lines[i].Trim();
                        if (line.StartsWith("DOWNLOAD LATENCY", StringComparison.OrdinalIgnoreCase))
                        {
                            // look ahead for IQM
                            for (int j = i + 1; j < Math.Min(i + 6, lines.Length); j++)
                            {
                                var t = lines[j].Trim();
                                if (t.StartsWith("IQM", StringComparison.OrdinalIgnoreCase) || t.StartsWith("IQM :", StringComparison.OrdinalIgnoreCase))
                                {
                                    var parts = t.Split(':');
                                    if (parts.Length >= 2 && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) dlIQM = v;
                                }
                            }
                        }
                        if (line.StartsWith("UPLOAD LATENCY", StringComparison.OrdinalIgnoreCase))
                        {
                            for (int j = i + 1; j < Math.Min(i + 6, lines.Length); j++)
                            {
                                var t = lines[j].Trim();
                                if (t.StartsWith("IQM", StringComparison.OrdinalIgnoreCase) || t.StartsWith("IQM :", StringComparison.OrdinalIgnoreCase))
                                {
                                    var parts = t.Split(':');
                                    if (parts.Length >= 2 && double.TryParse(parts[1].Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) ulIQM = v;
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            var internetIssues = new List<string>();
            if (speed.Download < DOWNLOAD_THRESHOLD) internetIssues.Add($"Download speed below threshold: {speed.Download} Mbps (< {DOWNLOAD_THRESHOLD} Mbps)");
            if (speed.Upload < UPLOAD_THRESHOLD) internetIssues.Add($"Upload speed below threshold: {speed.Upload} Mbps (< {UPLOAD_THRESHOLD} Mbps)");
            if (speed.Ping > PING_THRESHOLD) internetIssues.Add($"Ping is high: {speed.Ping} ms (> {PING_THRESHOLD} ms)");
            if (dlIQM.HasValue && dlIQM.Value >= DL_LATENCY_THRESHOLD) internetIssues.Add($"Download latency high (IQM): {dlIQM.Value} ms (>= {DL_LATENCY_THRESHOLD} ms)");
            if (ulIQM.HasValue && ulIQM.Value >= UL_LATENCY_THRESHOLD) internetIssues.Add($"Upload latency high (IQM): {ulIQM.Value} ms (>= {UL_LATENCY_THRESHOLD} ms)");

            if (internetIssues.Count == 0)
            {
                sb.AppendLine("✓ Internet connection appears healthy.");
            }
            else
            {
                sb.AppendLine("! Internet issues detected:");
                foreach (var it in internetIssues) sb.AppendLine("  - " + it);
            }

            // ISSUES SUMMARY: concise status for easy scanning
            try
            {
                sb.AppendLine();
                sb.AppendLine("ISSUES SUMMARY");
                sb.AppendLine(new string('=', 14));

                // CPU
                var cpuStatus = sys.CPUUsage >= CPU_THRESHOLD ? "WARNING" : "OK";
                sb.AppendLine($"CPU: {cpuStatus} ({Math.Round(sys.CPUUsage, 2)}%)");

                // RAM
                var ramStatus = sys.RAMUsage >= RAM_THRESHOLD ? "WARNING" : "OK";
                sb.AppendLine($"RAM: {ramStatus} ({Math.Round(sys.RAMUsage, 2)}%)");

                // Uptime
                var upDays = Math.Round(sys.Uptime.TotalDays, 2);
                var upStatus = sys.Uptime.TotalDays >= UPTIME_DAYS_THRESHOLD ? "WARNING" : "OK";
                sb.AppendLine($"Uptime: {upStatus} ({upDays} days)");
                if (upStatus == "WARNING") sb.AppendLine("  Suggestion: consider rebooting the system to reduce uptime and apply updates.");

                // Disk - compute worst (lowest) free percent across local fixed drives
                double worstFreePct = 100.0;
                string worstDrive = "N/A";
                try
                {
                    using var diskSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3");
                    foreach (ManagementObject dobj in diskSearcher.Get())
                    {
                        var drive = dobj["DeviceID"]?.ToString() ?? "?";
                        var sizeObj = dobj["Size"];
                        var freeObj = dobj["FreeSpace"];
                        if (sizeObj != null && freeObj != null)
                        {
                            var total = Convert.ToDouble(sizeObj);
                            var free = Convert.ToDouble(freeObj);
                            var freePct = total > 0 ? Math.Round((free / total) * 100.0, 2) : 0.0;
                            if (freePct < worstFreePct)
                            {
                                worstFreePct = freePct;
                                worstDrive = drive;
                            }
                        }
                    }
                }
                catch { }
                var diskStatus = worstFreePct < DISK_FREE_THRESHOLD ? "WARNING" : "OK";
                sb.AppendLine($"Disk: {diskStatus} ({worstDrive} {Math.Round(worstFreePct, 2)}% free)");

                // Download / Upload
                if (speed != null)
                {
                    var dlStatus = speed.Download < DOWNLOAD_THRESHOLD ? "WARNING" : "OK";
                    var ulStatus = speed.Upload < UPLOAD_THRESHOLD ? "WARNING" : "OK";
                    sb.AppendLine($"Download: {dlStatus} ({Math.Round(speed.Download, 2)} Mbps)");
                    sb.AppendLine($"Upload: {ulStatus} ({Math.Round(speed.Upload, 2)} Mbps)");

                    // Ping
                    var pingStatus = speed.Ping > PING_THRESHOLD ? "WARNING" : "OK";
                    sb.AppendLine($"Ping: {pingStatus} ({Math.Round(speed.Ping, 2)} ms)");

                    // IQM latencies (if available)
                    if (dlIQM.HasValue)
                    {
                        var dlLatStatus = dlIQM.Value >= DL_LATENCY_THRESHOLD ? "WARNING" : "OK";
                        sb.AppendLine($"Download IQM latency: {dlLatStatus} ({Math.Round(dlIQM.Value, 3)} ms)");
                    }
                    else sb.AppendLine("Download IQM latency: N/A");

                    if (ulIQM.HasValue)
                    {
                        var ulLatStatus = ulIQM.Value >= UL_LATENCY_THRESHOLD ? "WARNING" : "OK";
                        sb.AppendLine($"Upload IQM latency: {ulLatStatus} ({Math.Round(ulIQM.Value, 3)} ms)");
                    }
                    else sb.AppendLine("Upload IQM latency: N/A");
                }
                else
                {
                    sb.AppendLine("Download: N/A");
                    sb.AppendLine("Upload: N/A");
                    sb.AppendLine("Ping: N/A");
                    sb.AppendLine("Download IQM latency: N/A");
                    sb.AppendLine("Upload IQM latency: N/A");
                }

                sb.AppendLine();
            }
            catch { }

            // Try to include more detailed info from the human-readable speedtest results file if present
            try
            {
                var speedTxt = Path.Combine(Path.GetDirectoryName(outputFile) ?? string.Empty, "speedtest_results.txt");
                if (File.Exists(speedTxt))
                {
                    sb.AppendLine();
                    sb.AppendLine("SPEEDTEST DETAILS");
                    sb.AppendLine(new string('=', 16));
                    var lines = File.ReadAllLines(speedTxt);
                    var writtenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var l in lines)
                    {
                        var trimmed = l.Trim();
                        // Stop if the speedtest file continues into server connectivity or other large sections
                        if (trimmed.StartsWith("SERVER CONNECTIVITY CHECK", StringComparison.OrdinalIgnoreCase)) break;

                        // Skip decorative separators and the big header block
                        if (trimmed.StartsWith("*****") || trimmed.StartsWith("* SPEEDTEST RESULTS *") || trimmed.StartsWith("TEST TIME:", StringComparison.OrdinalIgnoreCase) == false && trimmed.StartsWith("* "))
                        {
                            // allow TEST TIME line, but skip other decorative lines starting with '*'
                            if (trimmed.StartsWith("TEST TIME:", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!writtenKeys.Contains("TestTime")) { sb.AppendLine(trimmed); writtenKeys.Add("TestTime"); }
                            }
                            continue;
                        }

                        // Skip sensitive/unwanted fields completely
                        if (trimmed.StartsWith("Host:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Server IP:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Server Port:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Server ID:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("MAC Address:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Is VPN:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Internal IP:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("External IP:", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            continue; // omit these
                        }

                        // Do NOT include the raw speed metrics here (they are already in INTERNET section)
                        if (trimmed.StartsWith("Download", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Upload", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Ping", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Jitter", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Packet Loss", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("SPEED RESULTS", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            continue;
                        }

                        // Allow test time, server name/location, ISP, latency blocks and result URL
                        if (trimmed.StartsWith("Name:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Location:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("ISP:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("Result URL:", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("DOWNLOAD LATENCY", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("UPLOAD LATENCY", StringComparison.OrdinalIgnoreCase)
                            || trimmed.TrimStart().StartsWith("Jitter", StringComparison.OrdinalIgnoreCase)
                            || trimmed.StartsWith("TEST TIME:", StringComparison.OrdinalIgnoreCase)
                            )
                        {
                            // dedupe repeated latency lines
                            var key = trimmed.Split(':')[0].Trim();
                            if (writtenKeys.Contains(key)) continue;
                            writtenKeys.Add(key);
                            sb.AppendLine(trimmed);
                        }
                        else
                        {
                            // ignore any other lines to avoid repeating data
                        }
                    }
                }
            }
            catch { }
        }
        sb.AppendLine();

        // SERVER CONNECTIVITY
    sb.AppendLine("SERVER CONNECTIVITY CHECK");
    sb.AppendLine(new string('=', 25));

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

    // Duplicate recor
    // 
    // d types removed. Use the public records declared in SystemInfo.cs

    // Backwards-compatible wrapper expected by AnalyzerService
    public static async Task<TcpingResult?> StartParallelTests(string target, string outputFolder, int testDurationSeconds)
    {
        var startMsg = "Starting Ace Cloud Gateway Servers Scanning";
        Console.WriteLine(startMsg);
        Log(startMsg);

        var list = new List<string> { target };
        var results = await RunConcurrentGatewayScans(list, outputFolder, testDurationSeconds);

        // After WinMTR/tcping sessions close, indicate we're generating reports
        var genMsg = "Generating Analysis Reports";
        Console.WriteLine(genMsg);
        Log(genMsg);

        var doneMsg = "Completed Ace Cloud Gateway Servers Scanning";
        Console.WriteLine(doneMsg);
        Log(doneMsg);

        return results.FirstOrDefault();
    }
}
}
