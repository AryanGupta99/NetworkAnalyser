using System;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace NativeAnalyzer
{
    class TestRunner
    {
        static async Task<int> Main(string[] args)
        {
            var outDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "Ace Network Result", "test_run");
            System.IO.Directory.CreateDirectory(outDir);
            var targets = new List<string> { "127.0.0.1", "8.8.8.8", "1.1.1.1" };
            var results = await Program.RunConcurrentGatewayScans(targets, outDir, 10);
            Console.WriteLine("RunConcurrentGatewayScans completed. Results:");
            foreach (var r in results) Console.WriteLine($"{r.Target}: {(string.IsNullOrEmpty(r.TCPingData) ? "no data" : r.TCPingData.Substring(0, Math.Min(80, r.TCPingData.Length)) )}");
            return 0;
        }
    }
}
