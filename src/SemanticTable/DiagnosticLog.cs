using System;
using System.IO;

namespace SemanticTable
{
    internal static class DiagnosticLog
    {
        private static readonly object Sync = new object();
        public static string PathName => Path.Combine(Path.GetTempPath(), "SemanticTable.log");

        public static void Write(string message)
        {
            try
            {
                lock (Sync)
                    File.AppendAllText(PathName,
                        DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff 'UTC'") + "  " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
