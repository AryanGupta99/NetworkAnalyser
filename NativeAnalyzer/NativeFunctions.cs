using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Threading.Tasks;
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
    public static class NativeFunctions
    {
        // Simplified helpers copied from Program.cs (non-UI entry points)
        public static void Log(string message, string? currentLogFile = null)
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
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
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

        public static ulong GetTotalPhysicalMemory()
        {
            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem)) return mem.ullTotalPhys;
            return 0;
        }

        public static ulong GetAvailableMemory()
        {
            var mem = new MEMORYSTATUSEX();
            if (GlobalMemoryStatusEx(mem)) return mem.ullAvailPhys;
            return 0;
        }

        public static SystemInfo GetSystemInfo(string outputFile)
        {
            // copy of Program.GetSystemInfo
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

        // Other helpers such as TakeScreenshot, StartSpeedTest, StartParallelTests, CreateSummaryReport, ParsePingTimes
        // For brevity, AnalyzerService will call into Program's versions if present; here provide small wrappers or stubs if necessary.
    }
}
