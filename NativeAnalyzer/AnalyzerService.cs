using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NativeAnalyzer
{
    public class AnalyzerService
    {
    public async Task<string> RunAnalysisAsync(string? baseFolder = null, IProgress<string>? progress = null, CancellationToken cancellationToken = default, IProgress<int>? percentProgress = null)
        {
            baseFolder ??= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Ace Network Result");
            var dateStamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var outputFolder = Path.Combine(baseFolder, dateStamp);
            Directory.CreateDirectory(outputFolder);

            progress?.Report($"Starting analysis. Output folder: {outputFolder}");

            try
            {
                var sys = NativeFunctions.GetSystemInfo(Path.Combine(outputFolder, "SystemInfo.txt"));
                progress?.Report("System information collected");

                var speed = await Program.StartSpeedTest(Path.Combine(outputFolder, "speedtest_results.txt"));
                progress?.Report("Speedtest completed");
                // give UI a short moment to display the speedtest completion before starting scans
                await Task.Delay(800, cancellationToken).ConfigureAwait(false);

                var targets = new[] { "RDGCHG.myrealdata.net", "RDGHTN.myrealdata.net", "RDGNV.myrealdata.net", "RDGATL.myrealdata.net", "RDGDEN.myrealdata.net" };
                progress?.Report("Starting Ace Cloud Gateway Servers Scanning");
                var tcpResults = await Program.RunConcurrentGatewayScans(targets.ToList(), outputFolder, 60, percentProgress, progress);
                // WinMTR windows may be closed by this point; show a report-generation status to the UI
                progress?.Report("Generating Analysis Reports");

                Program.CreateSummaryReport(sys, speed, tcpResults, Path.Combine(outputFolder, "summary_report.txt"));
                progress?.Report("Summary report created");
                progress?.Report("Completed Ace Cloud Gateway Servers Scanning");

                return outputFolder;
            }
            catch (OperationCanceledException)
            {
                progress?.Report("Analysis canceled");
                throw;
            }
            catch (Exception ex)
            {
                progress?.Report("Error: " + ex.Message);
                throw;
            }
        }
    }
}
