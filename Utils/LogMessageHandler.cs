namespace DingToolExcelTool.Utils
{
    using System.Windows;
    using DingToolExcelTool.Data;
    
    internal static class LogMessageHandler
    {
        public static MainWindow Window { get; private set; }

        public static void Init(MainWindow window) => Window = window;

        public static void AddInfo(string info)
        {
            Application.Current.Dispatcher.BeginInvoke(() =>
            {
                Window.AddLogItem(info, LogType.Info);
                Window.MoveLogListScroll();
            });
        }

        public static void AddWarn(string warn)
        {
            Application.Current.Dispatcher.BeginInvoke(() =>
            {
                Window.AddLogItem(warn, LogType.Warn);
                Window.MoveLogListScroll();
            });
        }

        public static void AddError(string error)
        {
            Application.Current.Dispatcher.BeginInvoke(() =>
            {
                Window.AddLogItem(error, LogType.Error);
                Window.MoveLogListScroll();
            });
        }
    }
}
