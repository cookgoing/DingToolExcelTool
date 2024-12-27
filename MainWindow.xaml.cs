namespace DingToolExcelTool
{
    using System.IO;
    using System.Diagnostics;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using Ookii.Dialogs.Wpf;
    using DingToolExcelTool.ExcelHandler;
    using DingToolExcelTool.Utils;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using static System.Runtime.InteropServices.JavaScript.JSType;

    public partial class MainWindow : Window
    {
        private enum ExcuteTypeEn
        { 
            ExportExcel,
            ClearOutputDir,
            RestoreDefaults,
        }

        private Dictionary<ScriptTypeEn, string> scriptTypeDic = new ()
        {
            {ScriptTypeEn.CSharp, "C#"},
        };

        private Dictionary<ExcuteTypeEn, string> excuteTypeDic = new()
        {
            {ExcuteTypeEn.ExportExcel, "导表"},
            {ExcuteTypeEn.ClearOutputDir, "清空缓存"},
            {ExcuteTypeEn.RestoreDefaults, "恢复默认值"},
        };

        private ExcuteTypeEn excutionType;
        private bool initedUI;

        public MainWindow()
        {
            initedUI = false;
            InitializeComponent();

            LogMessageHandler.Init(this);

            list_log.Items.Clear();
            cb_clientScriptType.Items.Clear();
            cb_serverScriptType.Items.Clear();
            cb_action.Items.Clear();

            foreach (ScriptTypeEn scriptType in Enum.GetValues<ScriptTypeEn>())
            {
                if (!scriptTypeDic.TryGetValue(scriptType, out string typeName))
                {
                    LogMessageHandler.AddError($"脚本类型，没有对应的文本显示字段：{scriptType}");
                    continue;
                }

                cb_clientScriptType.Items.Add(new ComboBoxItem { Content = typeName });
                cb_serverScriptType.Items.Add(new ComboBoxItem { Content = typeName });
            }

            foreach (ExcuteTypeEn excuteType in Enum.GetValues<ExcuteTypeEn>())
            {
                if (!excuteTypeDic.TryGetValue(excuteType, out string excutionName))
                {
                    LogMessageHandler.AddError($"操作类型，没有对应的文本显示字段：{excuteType}");
                    continue;
                }

                cb_action.Items.Add(new ComboBoxItem { Content = excutionName });
            }

            initedUI = true;
            RefreshUI();

            if (cb_action.Items[0] is ComboBoxItem actionType) actionType.IsSelected = true;
            excutionType = (ExcuteTypeEn)cb_action.SelectedIndex;
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            RefreshData();

            ExcelManager.Instance.Clear();
        }


        private void RefreshUI()
        {
            clientCheckBox.IsChecked = ExcelManager.Instance.Data.OutputClient;
            serverCheckBox.IsChecked = ExcelManager.Instance.Data.OutputServer;

            cb_clientScriptType.SelectedIndex = (int)ExcelManager.Instance.Data.ClientOutputInfo.ScriptType;
            cb_serverScriptType.SelectedIndex = (int)ExcelManager.Instance.Data.ServerOutputInfo.ScriptType;

            tb_excelPath.Text = ExcelManager.Instance.Data.ExcelInputRootDir;
            tb_clientPBMetaPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir;
            tb_clientPBScriptPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir;
            tb_clientPBDataPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir;
            tb_clientExcelScriptPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir;
            tb_clientErrorcodeFramePath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir;
            tb_clientErrorcodeBusinessPath.Text = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir;

            tb_serverPBMetaPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir;
            tb_serverPBScriptPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir;
            tb_serverPBDataPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir;
            tb_serverExcelScriptPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir;
            tb_serverErrorcodeFramePath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir;
            tb_serverErrorcodeBusinessPath.Text = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir;
        }

        private void RefreshData()
        {
            ExcelManager.Instance.Data.OutputClient = clientCheckBox.IsChecked ?? false;
            ExcelManager.Instance.Data.OutputServer = serverCheckBox.IsChecked ?? false;

            ExcelManager.Instance.Data.ClientOutputInfo.ScriptType = (ScriptTypeEn)cb_clientScriptType.SelectedIndex;
            ExcelManager.Instance.Data.ServerOutputInfo.ScriptType = (ScriptTypeEn)cb_serverScriptType.SelectedIndex;

            ExcelManager.Instance.Data.ExcelInputRootDir = tb_excelPath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir = tb_clientPBMetaPath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir = tb_clientPBScriptPath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir = tb_clientPBDataPath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir = tb_clientExcelScriptPath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir = tb_clientErrorcodeFramePath.Text;
            ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir = tb_clientErrorcodeBusinessPath.Text;

            ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir = tb_serverPBMetaPath.Text;
            ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir = tb_serverPBScriptPath.Text;
            ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir = tb_serverPBDataPath.Text;
            ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir = tb_serverExcelScriptPath.Text;
            ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir = tb_serverErrorcodeFramePath.Text;
            ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir = tb_serverErrorcodeBusinessPath.Text;
        }


        private void SelectorAction(TextBox pathTextor, string desc)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = desc,
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == true) pathTextor.Text = dialog.SelectedPath;
        }

        public int AddLogItem(string logStr, LogType logType)
        {
            return list_log.Items.Add(new ListBoxItem
            {
                Content = logStr,
                Foreground = logType == LogType.Error ? Brushes.Red : logType == LogType.Warn ? Brushes.Orange : Brushes.Black,
            });
        }

        public void MoveLogListScroll(int idx = -1)
        {
            if (idx == -1)
            {
                MoveLogListScroll(list_log.Items.Count - 1);
                return;
            }

            if (idx < 0 || idx >= list_log.Items.Count) return;

            list_log.ScrollIntoView(list_log.Items.GetItemAt(idx));
        }


        private void ExportExcel()
        {
            if (!ExcelManager.Instance.Data.OutputClient && !ExcelManager.Instance.Data.OutputServer)
            {
                LogMessageHandler.AddWarn($"没有输出路径，不进行相关的导出操作");
                return;
            }

            ClearOutputDir();

            ExcelManager.Instance.GenerateExcelHeadInfo();
            //ExcelManager.Instance.GenerateProtoMeta();
            //ExcelManager.Instance.GenerateProtoScript();
            //ExcelManager.Instance.GenerateProtoData();
            //ExcelManager.Instance.GenerateExcelScript();
        }

        private void ClearOutputDir()
        {
            if (ExcelManager.Instance.Data.OutputClient)
            {
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoMetaOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoScriptOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ProtoDataOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ClientOutputInfo.ExcelScriptOutputDir);

                string frameOutputDir = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir;
                string businessOutputDir = ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir;
                string fameFilePath = Path.Combine(frameOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);
                string businessFilePath = Path.Combine(businessOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);

                if (Path.Exists(fameFilePath)) File.Delete(fameFilePath);
                if (Path.Exists(businessFilePath)) File.Delete(businessFilePath);
            }

            if (ExcelManager.Instance.Data.OutputServer)
            {
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoMetaOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoScriptOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ProtoDataOutputDir);
                ExcelUtil.ClearDirectory(ExcelManager.Instance.Data.ServerOutputInfo.ExcelScriptOutputDir);

                string frameOutputDir = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir;
                string businessOutputDir = ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir;
                string fameFilePath = Path.Combine(frameOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);
                string businessFilePath = Path.Combine(businessOutputDir, SpecialExcelCfg.ErrorCodeFrameScriptFileName);

                if (Path.Exists(fameFilePath)) File.Delete(fameFilePath);
                if (Path.Exists(businessFilePath)) File.Delete(businessFilePath);
            }
        }

        private void RestoreDefault()
        {
            ExcelManager.Instance.ResetDefaultData();

            RefreshUI();
        }


        private void CheckBox_clientStateChanged(object sender, RoutedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.OutputClient = clientCheckBox.IsChecked ?? false; }
        private void CheckBox_serverStateChanged(object sender, RoutedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.OutputServer = serverCheckBox.IsChecked ?? false; }


        private void Btn_excelFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_excelPath, "选择表格文件夹");

        private void Btn_clientPBmetaFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientPBMetaPath, "选择客户端的proto meta的导出文件夹");
        private void Btn_clientPBScriptFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientPBScriptPath, "选择客户端的proto script的导出文件夹");
        private void Btn_clientPBDataFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientPBDataPath, "选择客户端的proto data的导出文件夹");
        private void Btn_clientExcelScriptFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientExcelScriptPath, "选择客户端的excel script的导出文件夹");
        private void Btn_clientECFrameFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientErrorcodeFramePath, "选择客户端的error code frame的导出文件夹");
        private void Btn_clientECBusinessFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_clientErrorcodeBusinessPath, "选择客户端的error code business的导出文件夹");

        private void Btn_serverPBMetaFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverPBMetaPath, "选择服务器的proto meta的导出文件夹");
        private void Btn_serverPBScriptFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverPBScriptPath, "选择服务器的proto script的导出文件夹");
        private void Btn_serverPBDataFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverPBDataPath, "选择服务器的proto data的导出文件夹");
        private void Btn_serverExcelScriptFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverExcelScriptPath, "选择服务器的excel script的导出文件夹");
        private void Btn_serverECFrameFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverErrorcodeFramePath, "选择服务器的error code frame的导出文件夹");
        private void Btn_serverECBusinessFolderSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_serverErrorcodeBusinessPath, "选择服务器的error code business的导出文件夹");

        private void Btn_preProcessSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_preProcessPath, "选择前处理的文件");
        private void Btn_aftProcessSelector(object sender, RoutedEventArgs e) => SelectorAction(tb_aftProcessPath, "选择后处理的文件");


        private void CB_clientScriptTypeChanged(object sender, SelectionChangedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.ClientOutputInfo.ScriptType = (ScriptTypeEn)cb_clientScriptType.SelectedIndex; }
        private void CB_serverScriptTypeChanged(object sender, SelectionChangedEventArgs e) { if (initedUI) ExcelManager.Instance.Data.ServerOutputInfo.ScriptType = (ScriptTypeEn)cb_serverScriptType.SelectedIndex; }
        private void CB_actionTypeChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!initedUI) return;

            excutionType = (ExcuteTypeEn)cb_action.SelectedIndex;
            grid_batch.Visibility = excutionType == ExcuteTypeEn.ExportExcel ? Visibility.Visible : Visibility.Collapsed;
        }


        private void Hyperlink_moreDetail(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.ToString(),
                UseShellExecute = true
            });

            e.Handled = true;
        }


        private void Btn_excute(object sender, RoutedEventArgs eventArg)
        {
            try
            {
                switch (excutionType)
                { 
                    case ExcuteTypeEn.ExportExcel:
                        if (!string.IsNullOrEmpty(tb_preProcessPath.Text)) ExcelUtil.ExcuteProgramFile(tb_preProcessPath.Text, tb_preProcessArgs.Text);

                        ExportExcel();

                        if (!string.IsNullOrEmpty(tb_aftProcessPath.Text)) ExcelUtil.ExcuteProgramFile(tb_preProcessPath.Text, tb_aftProcessArgs.Text);
                        LogMessageHandler.AddInfo("导表 完成");
                        break;
                    case ExcuteTypeEn.ClearOutputDir:
                        ClearOutputDir();
                        LogMessageHandler.AddInfo("清理缓存 完成");
                        break;
                    case ExcuteTypeEn.RestoreDefaults:
                        RestoreDefault();
                        LogMessageHandler.AddInfo("恢复默认值 完成");
                        break;
                }
            }
            catch (Exception e)
            {
                LogMessageHandler.AddError($"{e.Message}\n{e.StackTrace}");
            }
        }

        // todo: 文档
    }
}
