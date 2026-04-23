using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace MonTableurApp
{
    public partial class App : Application
    {
        private static readonly string CrashLogPath = Path.Combine(AppContext.BaseDirectory, "crash.log");

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            var window = new MainWindow();
            window.Show();
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogCrash("DispatcherUnhandledException", e.Exception);
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception exception)
            {
                LogCrash("AppDomainUnhandledException", exception);
            }
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogCrash("TaskSchedulerUnobservedTaskException", e.Exception);
        }

        private static void LogCrash(string source, Exception exception)
        {
            string message = $"""
[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}
{exception}

""";

            File.AppendAllText(CrashLogPath, message, new UTF8Encoding(false));
        }
    }
}
