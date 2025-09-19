using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shell;

namespace Ace_NDT_v2
{
    public partial class MainWindow : Window
    {
    private Process? nativeProcess;
    private CancellationTokenSource? _cts;
        private bool isAnalysisRunning = false;
        private bool isManualStop = false; // Flag to track if process was manually stopped
        private const string EmbeddedExeFile = "Ace_NDT_v2.tcping.exe";
        private const string EmbeddedDataFile = "Ace_NDT_v2.WinMTR.exe";
        private const string EmbeddedDataFile1 = "Ace_NDT_v2.anw.ps1";
        private const string EmbeddedDataFile2 = "Ace_NDT_v2.speedtest.exe";
            string tempDirectory = Path.Combine(Path.GetTempPath());
        //define temp dir for storage

        // Path to the results folder base
        private string aceNetworkResultPath;

        public MainWindow()
        {
            InitializeComponent();
            SetupWindowChrome();
            SetupTheme();

            // Create the temp directory if it doesn't exist
            if (!Directory.Exists(tempDirectory))
            {
                Directory.CreateDirectory(tempDirectory);
            }

            // Set up the base results folder path under Downloads
            aceNetworkResultPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                "Ace Network Result");

            // Set up the View Result button style
            ViewResultButton.IsEnabled = false;
            ViewResultButton.Background = new SolidColorBrush(Colors.LightGray);
        }


        private void vishal()
        {
            string lola = Guid.NewGuid().ToString();
        }

        private void SetupWindowChrome()
        {
            // Create window chrome for custom title bar
            WindowChrome.SetWindowChrome(this, new WindowChrome
            {
                CaptionHeight = 30,
                ResizeBorderThickness = new Thickness(5),
                CornerRadius = new CornerRadius(0),
                GlassFrameThickness = new Thickness(0),
                UseAeroCaptionButtons = true
            });
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void SetupTheme()
        {
            // Remove blur from MainGrid
            MainGrid.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(240, 240, 240, 240));

            // You could apply blur to a background element instead if desired
            // For example, if you had a decorative panel:
            // BackgroundPanel.Effect = new BlurEffect { Radius = 5, KernelType = KernelType.Gaussian };
        }

        private void StartAnalysis_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (isAnalysisRunning)
                    return;

                // Reset the manual stop flag when starting a new analysis
                isManualStop = false;
                // Disable the View Result button when starting a new analysis
                ViewResultButton.IsEnabled = false;
                ViewResultButton.Background = new SolidColorBrush(Colors.LightGray);

                // Extract the embedded resources (keep for backward compatibility)
                string exePath = null;
                string dataPath = null;
                string dataPathps = null;
                string dataPathspeedtest = null;
                // Skip extracting PowerShell script; we now always run the integrated C# analyzer

                // Copy native analyzer files into tempDirectory for isolated execution
                try
                {
                    string workspaceNative = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".", "NativeAnalyzer");
                    if (Directory.Exists(workspaceNative))
                    {
                        foreach (var f in Directory.GetFiles(workspaceNative, "*.*"))
                        {
                            var dest = Path.Combine(tempDirectory, Path.GetFileName(f));
                            try { File.Copy(f, dest, true); } catch { }
                        }
                    }
                }
                catch { }

                // Show disclaimer message before starting
                System.Windows.MessageBoxResult result = System.Windows.MessageBox.Show(
                    "This tool is designed exclusively for analyzing and troubleshooting Ace Cloud Hosting related network environments. The diagnostic process will require approximately 10 minutes to complete. " +
                    "Please make sure you close personal and confidential information from your monitor as this tool will take screenshots of your monitor.",
                    "ACE Network Analyser - Disclaimer",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Warning);

                // Only proceed if the user clicks OK
                if (result == MessageBoxResult.OK)
                {
                    try
                    {
                        // Run integrated C# analyzer (no PowerShell fallback) in background
                        isAnalysisRunning = true;

                        // Update UI
                        ProgressBorder.Visibility = Visibility.Visible;
                        ProgressBar.IsIndeterminate = true;
                        StartButton.IsEnabled = false;
                        StopButton.IsEnabled = true;
                        StatusText.Text = "Starting Ace Cloud Gateway Servers Scanning";

                        try
                        {
                            _cts = new CancellationTokenSource();
                            var svc = new NativeAnalyzer.AnalyzerService();
                            var progressReporter = new Progress<string>(s => Dispatcher.Invoke(() => StatusText.Text = s));
                            // Percent progress reporter (0-100)
                            var percentReporter = new Progress<int>(p => Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    ProgressBar.Value = Math.Max(0, Math.Min(100, p));
                                    // If p < 100, estimate seconds remaining using the test duration (60s)
                                    if (p >= 0 && p < 100)
                                    {
                                        int secondsRemaining = (int)Math.Ceiling((100 - p) * 60.0 / 100.0);
                                        SecondsRemainingText.Text = $"{secondsRemaining}s remaining";
                                    }
                                    else
                                    {
                                        SecondsRemainingText.Text = string.Empty;
                                    }
                                }
                                catch { }
                            }));

                            // Run analyzer in background without blocking UI
                            _ = Task.Run(async () =>
                            {
                                try
                                {
                                    // Pass user's Downloads folder as base output folder
                                    var downloadsBase = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Ace Network Result");
                                    var outFolder = await svc.RunAnalysisAsync(downloadsBase, progressReporter, _cts.Token, percentReporter);
                                    Dispatcher.Invoke(() =>
                                    {
                                        StatusText.Text = "Analysis completed.";
                                        ViewResultButton.IsEnabled = true;
                                        ViewResultButton.Background = new SolidColorBrush(Colors.LightGreen);
                                        isAnalysisRunning = false;
                                        ProgressBar.IsIndeterminate = false;
                                        ProgressBorder.Visibility = Visibility.Collapsed;
                                    });
                                }
                                catch (OperationCanceledException)
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        StatusText.Text = "Analysis canceled by user.";
                                        ProgressBar.IsIndeterminate = false;
                                        ProgressBorder.Visibility = Visibility.Collapsed;
                                    });
                                }
                                catch (Exception ex)
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        StatusText.Text = "Error: " + ex.Message;
                                        ProgressBar.IsIndeterminate = false;
                                        ProgressBorder.Visibility = Visibility.Collapsed;
                                    });
                                }
                                finally
                                {
                                    CleanupTempFiles();
                                }
                            });
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show($"Error starting integrated analyzer: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                            ResetUI();
                        }
                    }
                    catch (Exception ex)
                    {
                            System.Windows.MessageBox.Show($"Error starting analysis: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    // User canceled the operation
                    StatusText.Text = "Analysis canceled by user.";
                    CleanupTempFiles();
                }
            }
            catch (Exception ex)
            {
                    System.Windows.MessageBox.Show("Error Occured" + ex, "Error", MessageBoxButton.OK);
            }

        }

        private void Process_Exited(object sender, EventArgs e)
        {
            // Since this event is raised on a different thread, we need to invoke back to the UI thread
            Dispatcher.Invoke(() =>
            {
                ResetUI();

                // Only show completion message and enable view result button if the process wasn't manually stopped
                if (!isManualStop)
                {
                    MessageBox.Show("Analysis completed successfully!", "Network Analyzer - Analysis Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    StatusText.Text = "Analysis completed.";

                    // Enable the View Result button and change its color to green
                    ViewResultButton.IsEnabled = true;
                    ViewResultButton.Background = new SolidColorBrush(Colors.LightGreen);
                }

                // Cleanup the specific temp files after process completion
                CleanupTempFiles();
            });
        }

        private void StopAnalysis_Click(object sender, RoutedEventArgs e)
        {
            if (!isAnalysisRunning)
                return;

            try
            {
                // Set the flag to indicate this is a manual stop before killing the process
                isManualStop = true;

                // Cancel the integrated analyzer if running
                _cts?.Cancel();

                // Kill any native helper process if it exists
                try { if (nativeProcess != null && !nativeProcess.HasExited) nativeProcess.Kill(); } catch { }

                ResetUI();
                StatusText.Text = "Analysis interrupted.";

                // Clean up the temp files when the analysis is stopped
                CleanupTempFiles();

                // Show message box informing the user that the process was interrupted
                System.Windows.MessageBox.Show(
                    "The analysis was interrupted. The tool will close now.",
                    "Analysis Interrupted",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                // Close the application
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error stopping analysis: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ViewResult_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Get the most recently created folder in Ace Network Result directory
                string mostRecentFolder = GetMostRecentFolder();

                if (string.IsNullOrEmpty(mostRecentFolder))
                {
                    System.Windows.MessageBox.Show(
                        "No result folders found in Ace Network Result directory.",
                        "Folder Not Found",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                // Construct the full path to the summary report
                string summaryReportPath = Path.Combine(mostRecentFolder, "summary_report.txt");

                // Check if the results file exists
                if (File.Exists(summaryReportPath))
                {
                    // Open the file using the default associated application
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = summaryReportPath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    System.Windows.MessageBox.Show(
                        $"Results file not found at:\n{summaryReportPath}",
                        "File Not Found",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening results file: {ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // Clean up temp files before closing
            CleanupTempFiles();
            this.Close();
        }

        private void ResetUI()
        {
            isAnalysisRunning = false;
            ProgressBar.IsIndeterminate = false;
            StartButton.IsEnabled = true;
            StopButton.IsEnabled = false;
        }

        // Helper method to get the most recently created folder
        private string GetMostRecentFolder()
        {
            try
            {
                // Ensure the base directory exists
                if (!Directory.Exists(aceNetworkResultPath))
                {
                    return null;
                }

                // Get all subdirectories
                DirectoryInfo baseDir = new DirectoryInfo(aceNetworkResultPath);
                DirectoryInfo[] subDirs = baseDir.GetDirectories();

                // If there are no subdirectories, return null
                if (subDirs.Length == 0)
                {
                    return null;
                }

                // Find the most recently created directory
                DirectoryInfo mostRecent = subDirs.OrderByDescending(d => d.CreationTime).FirstOrDefault();

                return mostRecent?.FullName;
            }
            catch (Exception)
            {
                // In case of any error, return null
                return null;
            }
        }

        private string ExtractResourceToFile(string resourceName, string outputDirectory)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string fileName = Path.GetFileName(resourceName.Replace("Ace_NDT_v2.", ""));
            string outputPath = Path.Combine(outputDirectory, fileName);

            using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName))
            {
                if (resourceStream == null)
                {
                    throw new Exception($"Resource not found: {resourceName}");
                }

                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create))
                {
                    resourceStream.CopyTo(fileStream);
                }
            }

            return outputPath;
        }

        // Updated method to clean up specific files instead of entire directory
        private void CleanupTempFiles()
        {
            try
            {
                // Define specific files to delete from temp directory
                string[] filesToDelete = { "TCping.exe", "WinMTR.exe", "speedtest.exe" };

                foreach (string file in filesToDelete)
                {
                    string filePath = Path.Combine(tempDirectory, file);
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            File.Delete(filePath);
                            Debug.WriteLine($"Successfully deleted: {filePath}");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error deleting file {filePath}: {ex.Message}");
                            // Continue with next file even if this one fails
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't interrupt the application flow
                Debug.WriteLine($"Error during temp file cleanup: {ex.Message}");
            }
        }

        // We still need this method for backward compatibility in case some code calls it
        private void CleanupTempDirectory()
        {
            // Redirect to the new method
            CleanupTempFiles();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Make sure we clean up processes when closing app
            if (nativeProcess != null && !nativeProcess.HasExited)
            {
                try
                {
                    nativeProcess.Kill();
                }
                catch
                {
                    // Ignore errors on shutdown
                }
            }

            // Clean up the temporary files
            CleanupTempFiles();

            base.OnClosed(e);
        }
    }
}