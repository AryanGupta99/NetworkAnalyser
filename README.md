# Ace NDT - Ace Cloud Hosting Network Analyzer

## Overview

Ace NDT is a desktop application for Windows designed to help diagnose network connectivity issues for Ace Cloud Hosting customers. It provides a simple user interface to run a series of network diagnostic tests and collects the results in a local folder for analysis.

This project was originally based on PowerShell scripts but has been refactored into a native C#/.NET application to avoid restrictions on script execution and improve performance.

## Features

*   **Simple User Interface**: An intuitive WPF interface to start and stop the network analysis.
*   **Comprehensive Network Tests**:
    *   **System Information**: Collects basic system details (OS, CPU, RAM).
    *   **Speed Test**: Measures download and upload speeds.
    *   **TCP Ping**: Checks connectivity and latency to specified Ace Cloud Hosting servers.
    *   **MTR (My Traceroute)**: Traces the network path to the servers to identify potential packet loss or latency issues.
*   **Organized Results**: Saves all test outputs into a timestamped folder in the user's `Downloads\Ace Network Result` directory.
*   **Summary Report**: Generates a `summary_report.txt` file with a consolidated view of all test results.

## Project Structure

The solution is composed of two main projects:

*   `Ace NDT v2`: A C# WPF project that provides the main graphical user interface (GUI) for the application.
    *   `MainWindow.xaml`/`.cs`: The main application window and its logic.
*   `NativeAnalyzer`: A .NET console application that contains the core logic for performing the network analysis.
    *   `AnalyzerService.cs`: Orchestrates the different analysis steps.
    *   `Program.cs`: Contains helper functions to run external tools and generate reports.
    *   `NativeFunctions.cs`: Includes functions for gathering system information.

### External Tools

The `NativeAnalyzer` relies on the following external command-line tools, which are included in the repository:

*   `speedtest.exe`: For running internet speed tests.
*   `tcping.exe`: For performing TCP-based pings.
*   `WinMTR.exe`: For running MTR (My Traceroute) tests.

## Getting Started

### Prerequisites

*   .NET 8 SDK
*   Visual Studio 2022 (or another C# compatible IDE)

### Building the Project

1.  Clone the repository.
2.  Open the `Ace NDT v2.sln` file in Visual Studio.
3.  Build the solution (Build > Build Solution). This will restore NuGet packages and compile both projects.

### Running the Application

After a successful build, you can run the application in a few ways:

*   **From Visual Studio**: Set `Ace NDT v2` as the startup project and press F5 to start debugging.
*   **From the executable**: Navigate to the output directory (e.g., `bin\Debug\net8.0-windows`) and run `Ace NDT v2.exe`.

## How to Use

1.  Launch the application.
2.  Click the "Start Analysis" button.
3.  The application will begin running the network tests. The progress will be displayed in the output log on the main window.
4.  Once the analysis is complete, the "View Result" button will become active. Click it to open the folder containing the test results.
