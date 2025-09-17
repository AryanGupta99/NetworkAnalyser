NativeAnalyzer

A small console app port of the PowerShell `anw.ps1` script. It performs:
- Basic system info capture
- Optional Ookla Speedtest CLI invocation (if `speedtest.exe` is available)
- Optional `tcping.exe` and `WinMTR.exe` invocations if present in the same folder
- Generates a `summary_report.txt` and `SystemInfo.txt` in `~/Downloads/Ace Network Result/<timestamp>`

Build & run:

```powershell
cd "C:\Users\tannu\Desktop\working ps\NativeAnalyzer"
dotnet build
dotnet run --configuration Debug
```
