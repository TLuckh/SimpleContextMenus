namespace SimpleContextMenus;

using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

public static class GlobalErrorLogger
{
    private static string LogFilePath = "";
    public static void Initialize(string logFilePath)
    {
        if (LogFilePath != "")
            return;
        
        LogFilePath = Path.Combine(logFilePath, "SimpleContextMenus_error_log.log");

        AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
        {
            LogError(args.ExceptionObject as Exception, "UnhandledException");
        };

        TaskScheduler.UnobservedTaskException += (sender, args) =>
        {
            LogError(args.Exception, "UnobservedTaskException");
            args.SetObserved();
        };
    }

    private static void LogError(Exception? ex, string source)
    {
        string logMessage = $"[{DateTime.Now}] [{source}] {ex?.Message}{Environment.NewLine}{ex?.StackTrace}{Environment.NewLine}";
        File.AppendAllText(LogFilePath, logMessage);
    }
}