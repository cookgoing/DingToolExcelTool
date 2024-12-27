namespace DingToolExcelTool.ExcelHandler
{
    using System.IO;
    using Newtonsoft.Json;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Utils;

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
                MaxDegreeOfParallelism = Environment.ProcessorCount,
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
            Parallel.ForEach(excelPathArr, options, excelFilePath =>
            {
                try
                {
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
                    LogMessageHandler.AddError($"{e.Message}\n{e.StackTrace}");
                    cts.Cancel();
                }
            });

            return result;
        }

        public void GenerateProtoMeta()
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
        }
        
        public void GenerateProtoScript()
        {
            LogMessageHandler.AddInfo($"【生成proto原型文件】");

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
                string serverMetaOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                CommonExcelHandler.GenerateProtoScript(commonServerMetaInputFile, serverMetaOutputFile, serverScriptType);
                EnumExcelHandler.GenerateProtoScript(enumServerMetaInputFile, serverMetaOutputFile, serverScriptType);
                ErrorCodeExcelHandler.GenerateProtoScript(errorCodeServerMetaInputFile, serverMetaOutputFile, serverScriptType);
            }
        }
        
        public bool GenerateProtoData()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            Parallel.ForEach(excelPathArr, options, excelFilePath =>
            {
                try
                {
                    LogMessageHandler.AddInfo($"【生成proto数据】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                    if (turnonClient)
                    {
                        string clientProtoDataOutputDir = Data.ClientOutputInfo.ProtoDataOutputDir;
                        ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else CommonExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                    }
                    if (turnonServer)
                    {
                        string serverProtoDataOutputDir = Data.ServerOutputInfo.ProtoDataOutputDir;
                        ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else CommonExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                    }
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.AddError($"{e.Message}\n{e.StackTrace}");
                    cts.Cancel();
                }
            });

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

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else CommonExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                    }
                    if (turnonServer)
                    {
                        string serverExcelScriptOutputDir = Data.ServerOutputInfo.ExcelScriptOutputDir;
                        ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) EnumExcelHandler.GenerateProtoData(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) SingleExcelHandler.GenerateProtoData(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else CommonExcelHandler.GenerateProtoData(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    }
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.AddError($"{e.Message}\n{e.StackTrace}");
                    cts.Cancel();
                }
            });

            return result;
        }
    }
}
