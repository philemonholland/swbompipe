using System;
using System.Globalization;
using System.IO;

namespace BomPipePdmVaultInstaller;

internal static class BomPipeVaultInstallerLog
{
    private static readonly object SyncRoot = new object();

    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Error(string message, Exception exception)
    {
        Write("ERROR", string.Concat(message, Environment.NewLine, exception));
    }

    private static void Write(string level, string message)
    {
        try
        {
            var logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AFCA",
                "BOMPipe",
                "logs");
            Directory.CreateDirectory(logDirectory);

            var logPath = Path.Combine(logDirectory, "BomPipePdmVaultInstaller.log");
            var line = string.Format(
                CultureInfo.InvariantCulture,
                "[{0:O}] [{1}] {2}{3}",
                DateTimeOffset.Now,
                level,
                message,
                Environment.NewLine);

            lock (SyncRoot)
            {
                File.AppendAllText(logPath, line);
            }
        }
        catch
        {
            // Logging must never block install or uninstall.
        }
    }
}
