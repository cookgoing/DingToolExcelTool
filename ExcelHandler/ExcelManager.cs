namespace DingToolExcelTool.ExcelHandler
{
    using System.IO;
    using Newtonsoft.Json;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Utils;
    using DingToolExcelTool.ScriptHandler;
    
    internal class ExcelManager : Singleton<ExcelManager>
    {
        public DataWraper? Data { get; private set; }
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
                MaxDegreeOfParallelism = Environment.ProcessorCount - 1,
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

        public DataWraper? ReadCustomData(string path)
        {
            using StreamReader sr = new StreamReader(path);
            return new (JsonConvert.DeserializeObject<CustomDataInfo>(sr.ReadToEnd()));
        }

        public void WriteCustomData()
        {
            string dataDir = Path.GetDirectoryName(GeneralCfg.CustomDataPath);
            if (!Directory.Exists(dataDir)) Directory.CreateDirectory(dataDir);

            using StreamWriter sw = new StreamWriter(GeneralCfg.CustomDataPath);
            sw.Write(JsonConvert.SerializeObject(Data?.Data));
        }

        public void ResetDefaultData()
        {
            Data = ReadCustomData(GeneralCfg.DefaultDataPath);
            WriteCustomData();
        }

        public async Task<bool> GenerateExcelHeadInfo()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            string enumExcelPath = Array.Find(excelPathArr, Path => Path.EndsWith("Enum.xlsx"));
            if (!string.IsNullOrEmpty(enumExcelPath))
            {
                LogMessageHandler.AddInfo($"【解析表头】:{enumExcelPath}");
                await Task.Run(() => EnumExcelHandler.GenerateExcelHeadInfo(enumExcelPath));
            }

            await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
            {
                try
                {
                    if (excelFilePath == enumExcelPath) return;
                    if (token.IsCancellationRequested) return;

                    LogMessageHandler.AddInfo($"【解析表头】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                    if (excelFileName == SpecialExcelCfg.EnumExcelName) await EnumExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ErrorCodeExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await SingleExcelHandler.GenerateExcelHeadInfo(excelFilePath);
                    else await CommonExcelHandler.GenerateExcelHeadInfo(excelFilePath);
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

        public async Task<bool> GenerateProtoMeta()
        {
            LogMessageHandler.AddInfo($"【生成proto原型文件】");
            
            async Task<bool> DoFunc(string commonOuput, string errorCodeOutput, string enumOutput, bool isClient)
            {
                try
                {
                    await Task.WhenAll(CommonExcelHandler.GenerateProtoMeta(commonOuput, isClient)
                            , ErrorCodeExcelHandler.GenerateProtoMeta(errorCodeOutput, isClient)
                            , EnumExcelHandler.GenerateProtoMeta(enumOutput, isClient));
                }
                catch (Exception e)
                {
                    
                    LogMessageHandler.LogException(e);
                    return false;
                }

                return true;
            }

            bool result = true;
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;
            if (turnonClient)
            {
                string commonClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeClientMetaOutputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");

                result &= await DoFunc(commonClientMetaOutputFile, errorCodeClientMetaOutputFile, enumClientMetaOutputFile, true);
            }
            if (turnonServer)
            {
                string commonServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeServerMetaOutputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");

                result &= await DoFunc(commonServerMetaOutputFile, errorCodeServerMetaOutputFile, enumServerMetaOutputFile, false);
            }

            return result;
        }
        
        public async Task<bool> GenerateProtoScript()
        {
            LogMessageHandler.AddInfo($"【生成proto脚本文件】");

            async Task<bool> DoFunc(string commonInput, string errorCodeInput, string enumInput, string output, ScriptTypeEn scriptType)
            {
                try
                {
                    await Task.WhenAll(CommonExcelHandler.GenerateProtoScript(commonInput, output, scriptType)
                            , ErrorCodeExcelHandler.GenerateProtoScript(errorCodeInput, output, scriptType)
                            , EnumExcelHandler.GenerateProtoScript(enumInput, output, scriptType));
                }
                catch (Exception e)
                {

                    LogMessageHandler.LogException(e);
                    return false;
                }

                return true;
            }

            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            if (turnonClient)
            {
                string commonClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeClientMetaInputFile = Path.Combine(Data.ClientOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string clientProtoScriptOutputDir = Data.ClientOutputInfo.ProtoScriptOutputDir;
                ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                await DoFunc(commonClientMetaInputFile, errorCodeClientMetaInputFile, enumClientMetaInputFile, clientProtoScriptOutputDir, clientScriptType);
            }
            if (turnonServer)
            {
                string commonServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{CommonExcelCfg.ProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string enumServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string errorCodeServerMetaInputFile = Path.Combine(Data.ServerOutputInfo.ProtoMetaOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}");
                string serverProtoScriptOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                await DoFunc(commonServerMetaInputFile, errorCodeServerMetaInputFile, enumServerMetaInputFile, serverProtoScriptOutputFile, serverScriptType);
            }

            return true;
        }
        
        public async Task<bool> GenerateProtoData()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            if (turnonClient)
            {
                string clientProtoDataOutputDir = Data.ClientOutputInfo.ProtoDataOutputDir;
                ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                await Task.Run(async () =>
                {
                    LogMessageHandler.AddInfo($"【客户端代码动态编译】");
                    string clientProtoScriptOutputDir = Data.ClientOutputInfo.ProtoScriptOutputDir;
                    IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(clientScriptType);
                    string[] scriptPathArr = Directory.GetFiles(clientProtoScriptOutputDir, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                    string[] scriptContent = new string[scriptPathArr.Length];
                    for (int i = 0; i < scriptPathArr.Length; ++i)
                    {
                        scriptContent[i] = await File.ReadAllTextAsync(scriptPathArr[i]);
                    }
                    scriptHandler.DynamicCompile(scriptContent);
                });

                await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
                {
                    try
                    {
                        if (token.IsCancellationRequested) return;

                        LogMessageHandler.AddInfo($"【客户端生成proto数据】:{excelFilePath}");
                        string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                        if (excelFileName == SpecialExcelCfg.EnumExcelName) await EnumExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await SingleExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
                        else await CommonExcelHandler.GenerateProtoData(excelFilePath, clientProtoDataOutputDir, true, clientScriptType);
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

                await Task.Run(async () =>
                {
                    LogMessageHandler.AddInfo($"【服务器代码动态编译】");
                    string serverProtoScriptOutputFile = Data.ServerOutputInfo.ProtoScriptOutputDir;
                    IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(serverScriptType);
                    string[] scriptPathArr = Directory.GetFiles(serverProtoScriptOutputFile, $"*{scriptHandler.Suffix}", SearchOption.AllDirectories);
                    string[] scriptContent = new string[scriptPathArr.Length];
                    for (int i = 0; i < scriptPathArr.Length; ++i)
                    {
                        scriptContent[i] = await File.ReadAllTextAsync(scriptPathArr[i]);
                    }
                    scriptHandler.DynamicCompile(scriptContent);
                });

                await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
                {
                    try
                    {
                        if (token.IsCancellationRequested) return;

                        LogMessageHandler.AddInfo($"【服务器生成proto数据】:{excelFilePath}");
                        string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);
                        if (excelFileName == SpecialExcelCfg.EnumExcelName) await EnumExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ErrorCodeExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await SingleExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
                        else await CommonExcelHandler.GenerateProtoData(excelFilePath, serverProtoDataOutputDir, false, serverScriptType);
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

        public async Task<bool> GenerateExcelScript()
        {
            bool result = true;
            string[] excelPathArr = Directory.GetFiles(Data.ExcelInputRootDir, "*.xlsx", SearchOption.AllDirectories);
            bool turnonClient = Data.OutputClient;
            bool turnonServer = Data.OutputServer;

            await Parallel.ForEachAsync(excelPathArr, options, async (excelFilePath, token) =>
            {
                try
                {
                    if (token.IsCancellationRequested) return;

                    LogMessageHandler.AddInfo($"【生成Excel 脚本】:{excelFilePath}");
                    string excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                    if (turnonClient)
                    {
                        string clientExcelScriptOutputDir = Data.ClientOutputInfo.ExcelScriptOutputDir;
                        ScriptTypeEn clientScriptType = Data.ClientOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) await EnumExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await SingleExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                        else await CommonExcelHandler.GenerateExcelScript(excelFilePath, clientExcelScriptOutputDir, true, clientScriptType);
                    }
                    if (turnonServer)
                    {
                        string serverExcelScriptOutputDir = Data.ServerOutputInfo.ExcelScriptOutputDir;
                        ScriptTypeEn serverScriptType = Data.ServerOutputInfo.ScriptType;

                        if (excelFileName == SpecialExcelCfg.EnumExcelName) await EnumExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName == SpecialExcelCfg.ErrorCodeExcelName) await ErrorCodeExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else if (excelFileName.StartsWith(SpecialExcelCfg.SingleExcelPrefix)) await SingleExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                        else await CommonExcelHandler.GenerateExcelScript(excelFilePath, serverExcelScriptOutputDir, false, serverScriptType);
                    }
                }
                catch (Exception e)
                {
                    result = false;
                    LogMessageHandler.LogException(e);
                    cts.Cancel();
                }
            });

            await Task.CompletedTask;
            return result;
        }
    }
}
