namespace DingToolExcelTool.ExcelHandler
{
    using System.IO;
    using Newtonsoft.Json;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Utils;
    using DingToolExcelTool.ScriptHandler;
    using System.Runtime.InteropServices;

    internal class ExcelManager : Singleton<ExcelManager>
    {
        public CustomDataInfo? Data { get; private set; }
        public CommonExcelHandler CommonExcelHandler { get; private set; }
        public EnumExcelHandler EnumExcelHandler { get; private set; }
        public ErrorCodeExcelHandler ErrorCodeExcelHandler { get; private set; }
        public SingleExcelHandler SingleExcelHandler { get; private set; }

        private CancellationTokenSource cts;
        private ParallelOptions options;

        public ExcelManager()
        {
            Data = ReadCustomData(GeneralCfg.CustomDataPath);
            CommonExcelHandler = new();
            EnumExcelHandler = new();
            ErrorCodeExcelHandler = new();
            SingleExcelHandler = new();

            cts = new CancellationTokenSource();
            options = new ParallelOptions
            {
                //MaxDegreeOfParallelism = Environment.ProcessorCount,
                CancellationToken = cts.Token
            };
            
            CommonExcelHandler.Init();
            EnumExcelHandler.Init();
            ErrorCodeExcelHandler.Init();
            SingleExcelHandler.Init();
        }

        public void Clear()
        {
            cts.Dispose();
            WriteCustomData();
            
            CommonExcelHandler.Clear();
            EnumExcelHandler.Clear();
            ErrorCodeExcelHandler.Clear();
            SingleExcelHandler.Clear();
        }

        public void Reset()
        {
            cts = new CancellationTokenSource();
            options.CancellationToken = cts.Token;
            
            CommonExcelHandler.Clear();
            EnumExcelHandler.Clear();
            ErrorCodeExcelHandler.Clear();
            SingleExcelHandler.Clear();
        }

        public CustomDataInfo? ReadCustomData(string path)
        {
            using StreamReader sr = new StreamReader(path);
            return JsonConvert.DeserializeObject<CustomDataInfo>(sr.ReadToEnd());
        }

        public void WriteCustomData()
        {
            string dataDir = Path.GetDirectoryName(GeneralCfg.CustomDataPath);
            if (!Directory.Exists(dataDir)) Directory.CreateDirectory(dataDir);

            using StreamWriter sw = new StreamWriter(GeneralCfg.CustomDataPath);
            sw.Write(JsonConvert.SerializeObject(Data));
        }

        public void ResetDefaultData()
        {
            Data = ReadCustomData(GeneralCfg.DefaultDataPath);
            WriteCustomData();
        }

        public bool GenerateExcelHeadInfo()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            string enumExcelPath = Array.Find(excelPathArr, Path => Path.EndsWith("Enum.xlsx"));
            if (!string.IsNullOrEmpty(enumExcelPath))
            {
                LogMessageHandler.AddInfo($"【解析表头】:{enumExcelPath}");
                EnumExcelHandler.GenerateExcelHeadInfo(enumExcelPath);
            }

            Parallel.ForEach(excelPathArr, options, excelFilePath =>
            {
                try
                {
                    if (excelFilePath == enumExcelPath) return;

                    LogMessageHandler.AddInfo($"【解析表头】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else CommonExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });

            return result;
        }

        public bool GenerateProtoMeta()
        {
            LogMessageHandler.AddInfo($"【生成proto原型文件】");
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            if (turnonClient)
            {
                string commonClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");

                CommonExcelHandler.GenerateProtoMeta(commonClientMetaOutputFile, true);
                EnumExcelHandler.GenerateProtoMeta(enumClientMetaOutputFile, true);
                ErrorCodeExcelHandler.GenerateProtoMeta(errorCodeClientMetaOutputFile, true);
            }
            if (turnonServer)
            {
                string commonServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");

                CommonExcelHandler.GenerateProtoMeta(commonServerMetaOutputFile, false);
                EnumExcelHandler.GenerateProtoMeta(enumServerMetaOutputFile, false);
                ErrorCodeExcelHandler.GenerateProtoMeta(errorCodeServerMetaOutputFile, false);
            }

            return true;
        }
        
        public bool GenerateProtoScript()
        {
            LogMessageHandler.AddInfo($"【生成proto脚本文件】");

            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            if (turnonClient)
            {
                string commonClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string clientProtoScriptOutputDir = Data.ClientOutputInfo.ProtoScriptOutputDir;
                ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                CommonExcelHandler.GenerateProtoScript(commonClientMetaInputFile, clientProtoScriptOutputDir, clientScriptType);
                EnumExcelHandler.GenerateProtoScript(enumClientMetaInputFile, clientProtoScriptOutputDir, clientScriptType);
                ErrorCodeExcelHandler.GenerateProtoScript(errorCodeClientMetaInputFile, clientProtoScriptOutputDir, clientScriptType);
            }
            if (turnonServer)
            {
                string commonServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string serverProtoScriptOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                CommonExcelHandler.GenerateProtoScript(commonServerMetaInputFile, serverProtoScriptOutputFile, serverScriptType);
                EnumExcelHandler.GenerateProtoScript(enumServerMetaInputFile, serverProtoScriptOutputFile, serverScriptType);
                ErrorCodeExcelHandler.GenerateProtoScript(errorCodeServerMetaInputFile, serverProtoScriptOutputFile, serverScriptType);
            }

            return true;
        }
        
        public bool GenerateProtoData()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            if (turnonClient)
            {
                string clientProtoDataOutputDir = Data.ClientOutputInfo.ProtoDataOutputDir;
                ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                LogMessageHandler.AddWarn($"【客户端代码动态编译】");
                string clientProtoScriptOutputDir = Data.ClientOutputInfo.ProtoScriptOutputDir;
                IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(clientScriptType);
                string[] scriptPathArr = Directory.GetFiles(clientProtoScriptOutputDir, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                string[] scriptContent = new string[scriptPathArr.Length];
                for (int i = 0; i < scriptPathArr.Length; ++i)
                {
                    scriptContent[i] = File.ReadAllText(scriptPathArr[i]);
                }
                scriptHandler.DynamicCompile(scriptContent);

                Parallel.ForEach(excelPathArr, options, async excelFilePath =>
                {
                    try
                    {
                        LogMessageHandler.AddInfo($"【客户端生成proto数据】:{excelFilePath}");
                        string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else CommonExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                    }
                    catch (Exception e)
                    {
                        result = false;
                        LogMessageHandler.LogException(e);
                        cts.Cancel();
                    }
                });
            }

            if (!result) return false;

            if (turnonServer)
            {
                string serverProtoDataOutputDir = Data.ServerOutputInfo.ProtoDataOutputDir;
                ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                LogMessageHandler.AddWarn($"【服务器代码动态编译】");
                string serverProtoScriptOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(serverScriptType);
                string[] scriptPathArr = Directory.GetFiles(serverProtoScriptOutputFile, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                string[] scriptContent = new string[scriptPathArr.Length];
                for (int i = 0; i < scriptPathArr.Length; ++i)
                {
                    scriptContent[i] = File.ReadAllText(scriptPathArr[i]);
                }
                scriptHandler.DynamicCompile(scriptContent);

                Parallel.ForEach(excelPathArr, options, async excelFilePath =>
                {
                    try
                    {
                        LogMessageHandler.AddInfo($"【服务器生成proto数据】:{excelFilePath}");
                        string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else CommonExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                    }
                    catch (Exception e)
                    {
                        result = false;
                        LogMessageHandler.LogException(e);
                        cts.Cancel();
                    }
                });
            }

            return result;
        }

        public bool GenerateExcelScript()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            Parallel.ForEach(excelPathArr, options, excelFilePath =>
            {
                try
                {
                    LogMessageHandler.AddInfo($"【生成Excel 脚本】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                    if (turnonClient)
                    {
                        string clientExcelScriptOutputDir = Data.ClientOutputInfo.ExcelScriptOutputDir;
                        ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else CommonExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                    }
                    if (turnonServer)
                    {
                        string serverExcelScriptOutputDir = Data.ServerOutputInfo.ExcelScriptOutputDir;
                        ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else CommonExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    }
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });

            return result;
        }
    }
}
