using System.Globalization;

namespace SolidWorksBOMAddin;

internal static class BomPipeLog
{
    private static readonly object SyncRoot = new();

    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Error(string message, Exception exception)
    {
        Write("ERROR", $"{message}{Environment.NewLine}{exception}");
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

            var logPath = Path.Combine(logDirectory, "SolidWorksBOMAddin.log");
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
            // Logging must never break the add-in host.
        }
    }
}
