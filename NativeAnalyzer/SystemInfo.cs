using System;

namespace NativeAnalyzer
{
    public record SystemInfo
    {
        public double RAMUsage { get; init; }
        public double CPUUsage { get; init; }
        public TimeSpan Uptime { get; init; }
    }

    public record SpeedTestResult
    {
        public double Download { get; init; }
        public double Upload { get; init; }
        public double Ping { get; init; }
        public double Jitter { get; init; }
        public double PacketLoss { get; init; }
    }

    public record TcpingResult
    {
        public string Target { get; init; } = string.Empty;
        public string TCPingData { get; init; } = string.Empty;
    }
}