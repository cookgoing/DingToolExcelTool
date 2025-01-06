namespace DingToolExcelTool.ExcelHandler
{
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using ClosedXML.Excel;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.ScriptHandler;
    using DingToolExcelTool.Utils;
    using DocumentFormat.OpenXml.Spreadsheet;

    internal class ErrorCodeExcelHandler : CommonExcelHandler
    {
        public ConcurrentDictionary<string, ErrorCodeScriptInfo> ErrorCodeDic { get; private set; }//<script name, ErrorCodeScriptInfo>

        public override void Init() 
        {
            base.Init();
            ErrorCodeDic = new(); 
        }

        public override void Clear() 
        {
            base.Clear();
            ErrorCodeDic?.Clear();
        } 

        public override async Task GenerateExcelHeadInfo(string excelInputFile)
        {
            if (ErrorCodeDic == null) throw new Exception("[GenerateExcelHeadInfo] ErrorCodeDic == null");
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelHeadInfo] 表路径不存在：{excelInputFile}");

            await base.GenerateExcelHeadInfo(excelInputFile);

            string excelFileName = Path.GetFileNameWithoutExtension(excelInputFile);
            using XLWorkbook wb = new (excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                string headMessageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
                if (!HeadInfoDic.TryGetValue(headMessageName, out ExcelHeadInfo headInfo)) throw new Exception($"没有[{headMessageName}] 的表头信息");

                List<string> errorCodeFieldNames = new(SpecialExcelCfg.ErrorCodeFixedField.Keys);
                headInfo.Fields.ForEach(fieldInfo =>
                {
                    if (SpecialExcelCfg.ErrorCodeFixedField.TryGetValue(fieldInfo.Name, out string typeName))
                    {
                        if (fieldInfo.LocalizationImg || fieldInfo.LocalizationTxt) typeName = ExcelUtil.ClearTypeSymbol(typeName);
                        if (typeName != fieldInfo.Type) throw new Exception($"[GenerateExcelHeadInfo]. ErrorCode表中有一个字段的类型不一致：{fieldInfo.Name} - {fieldInfo.Type}; 类型应该是：{typeName}");

                        errorCodeFieldNames.Remove(fieldInfo.Name);
                    }
                });

                StringBuilder sb = new();
                foreach (string field in errorCodeFieldNames) sb.Append(field).Append(',');
                if (errorCodeFieldNames.Count > 0)
                {
                    sb.Remove(sb.Length - 1, 1);
                    throw new Exception($"[GenerateExcelHeadInfo] ErrorCode表[{headMessageName}]有缺失的字段: {sb}");
                }

                ErrorCodeScriptInfo scriptInfo = new();
                scriptInfo.SheetName = sheet.Name;
                scriptInfo.Fields = new(10);
                HashSet<string> codeStrSet = new();
                HashSet<int> codeSet = new();
                foreach (IXLRow row in sheet.RowsUsed())
                {
                    string typeName = row.Cell(1).GetString();
                    bool isTpyeRow = typeName.StartsWith('#');
                    if (isTpyeRow) continue;

                    ErrorCodeScriptFieldInfo errorCodeField = new();
                    scriptInfo.Fields.Add(errorCodeField);
                    foreach (IXLCell cell in row.Cells(false))
                    {
                        int columnIdx = cell.Address.ColumnNumber;
                        if (columnIdx == 1) continue;
                        if (ExcelUtil.IsRearMergedCell(cell)) continue;

                        
                        int fieldIdx = headInfo.GetFieldIdx(columnIdx);
                        if (fieldIdx == -1) throw new Exception($"[GenerateExcelHeadInfo] 表：{headMessageName} 存在字段没有和类型关联上. Address: {cell.Address}");

                        ExcelFieldInfo fieldInfo = headInfo.Fields[fieldIdx];
                        string columnContent = cell.GetString().Trim();
                        switch (fieldInfo.Name.ToLower())
                        {
                            case "code":
                                if (!int.TryParse(columnContent, out int codeValue)) throw new Exception($"[GenerateExcelHeadInfo]. 表：{headMessageName}, address: {cell.Address} code[{codeValue}]不能解析成整形");
                                errorCodeField.Code = codeValue;
                                break;
                            case "codestr":
                                if (!codeStrSet.Add(columnContent)) throw new Exception($"[GenerateExcelHeadInfo]. 表：{headMessageName} 中出现了相同名字的 code 名字：{columnContent}; Address: {cell.Address}");

                                errorCodeField.CodeStr = columnContent;
                                break;
                            case "comment":
                                errorCodeField.Comment = columnContent;
                                break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(scriptInfo.SheetName) && scriptInfo.Fields?.Count > 0) ErrorCodeDic.TryAdd(scriptInfo.SheetName, scriptInfo);
            }
        }

        public override async Task GenerateProtoMeta(string metaOutputFile, bool isClient)
        {
            LogMessageHandler.AddInfo($"[ErrorCodeExcelHandler][GenerateProtoMeta]: {metaOutputFile}, isClient: {isClient}");
            if (string.IsNullOrEmpty(metaOutputFile))
            {
                LogMessageHandler.AddWarn($"[GenerateProtoMeta] 不存在 proto meta 的输出路径，将不会执行输出操作");
                return;
            }

            string dirPath = Path.GetDirectoryName(metaOutputFile);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
            using StreamWriter metaSW = new(metaOutputFile);
            StringBuilder metaWriter = new();

            metaWriter.AppendLine(@$"
syntax = ""proto3"";
package {GeneralCfg.ProtoMetaPackageName};
import ""{SpecialExcelCfg.EnumProtoMetaFileName}{GeneralCfg.ProtoMetaFileSuffix}"";").AppendLine();

            ExcelHeadInfo headInfo = HeadInfoDic.First().Value;
            metaWriter.AppendLine($"message {SpecialExcelCfg.ErrorCodeProtoMessageName}").AppendLine("{");
            int messageFieldIdx = 0;
            foreach (ExcelFieldInfo fieldInfo in headInfo.Fields)
            {
                if ((platform & fieldInfo.Platform) == 0) continue;

                metaWriter.Append($"\t{ExcelUtil.ToProtoType(fieldInfo.Type)} {fieldInfo.Name} = {++messageFieldIdx};");
                if (string.IsNullOrEmpty(fieldInfo.Comment)) metaWriter.AppendLine();
                else metaWriter.AppendLine($"//{fieldInfo.Comment}");
            }
            metaWriter.AppendLine("}").AppendLine();

            await metaSW.WriteAsync(metaWriter);
            metaSW.Close();
            metaWriter.Clear();


            string metaFileName = Path.GetFileName(metaOutputFile);
            string metaName = metaFileName.Replace(GeneralCfg.ProtoMetaFileSuffix, string.Empty);
            string listFilePath = Path.Combine(dirPath, $"{metaName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}{GeneralCfg.ProtoMetaFileSuffix}");
            using StreamWriter listSW = new(listFilePath);
            StringBuilder listWriter = new();

            listWriter.AppendLine(@$"
syntax = ""proto3"";
package {GeneralCfg.ProtoMetaPackageName};
import ""{metaFileName}"";").AppendLine();


            listWriter.AppendLine(@$"message {SpecialExcelCfg.ErrorCodeProtoMessageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}
{{
    repeated {SpecialExcelCfg.ErrorCodeProtoMessageName} {CommonExcelCfg.ProtoMetaListFieldName} = 1;
}}").AppendLine();

            await listSW.WriteAsync(listWriter);
            listSW.Close();
            listWriter.Clear();
        }

        public override async Task GenerateProtoData(string excelInputFile, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType)
        {
            if (!File.Exists(excelInputFile)) throw new Exception($"[ErrorCodeExcelHandler.GenerateProtoData] 表路径不存在：{excelInputFile}");
            if (string.IsNullOrEmpty(protoDataOutputDir))
            {
                LogMessageHandler.AddWarn($"[ErrorCodeExcelHandler.GenerateProtoData] 不存在 proto data 的输出路径，将不会执行输出操作");
                return;
            }

            IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
            PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
            string excelFileName = Path.GetFileNameWithoutExtension(excelInputFile);
            using XLWorkbook wb = new(excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
                if (!HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) throw new Exception($"[GenerateProtoData] 表: {messageName} 没有 headInfo");

                AddProtoDataInScriptObj(scriptHandler, platform, sheet, SpecialExcelCfg.ErrorCodeProtoMessageName, headInfo);
            }

            string protoDataOutputFile = Path.Combine(protoDataOutputDir, $"{SpecialExcelCfg.ErrorCodeProtoMessageName}{GeneralCfg.ProtoDataFileSuffix}");
            scriptHandler.SerializeObjInProto($"{SpecialExcelCfg.ErrorCodeProtoMessageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", protoDataOutputFile);

            await Task.CompletedTask;
        }

        public override async Task GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType) 
        {
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelScript] 表路径不存在：{excelInputFile}");
            if (string.IsNullOrEmpty(excelScriptOutputDir))
            {
                LogMessageHandler.AddWarn($"[GenerateExcelScript] 不存在 Excel Script 的输出路径，将不会执行输出操作");
                return;
            }

            IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
            string excelRelativePath = ExcelUtil.GetExcelRelativePath(excelInputFile);
            string excelRelativeDir = Path.GetDirectoryName(excelRelativePath);
            using XLWorkbook wb = new(excelInputFile);
            int sheetCount = wb.Worksheets.Count;
            ExcelHeadInfo headInfo = HeadInfoDic.First().Value;
            string messageName = SpecialExcelCfg.ErrorCodeProtoMessageName;
            string outputFilePath = Path.Combine(excelScriptOutputDir, excelRelativeDir ?? "", $"{messageName}{GeneralCfg.ExcelScriptFileSuffix(scriptType)}");

            await scriptHandler.GenerateExcelScript(headInfo, messageName, outputFilePath, isClient);


            string frameOutputDir = isClient ? ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeFrameDir
                                                : ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeFrameDir;
            string businessOutputDir = isClient ? ExcelManager.Instance.Data.ClientOutputInfo.ErrorCodeBusinessDir
                                                : ExcelManager.Instance.Data.ServerOutputInfo.ErrorCodeBusinessDir;

            if (frameOutputDir == null || businessOutputDir == null) throw new Exception($"[GenerateExcelScript]. ErrorCode 需要有两个导出路径。 {frameOutputDir} | {businessOutputDir}");

            IScriptExcelHandler scriptExcelHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
            IScriptSpecialExcelHandler scriptSpecialHandler = ExcelUtil.GetScriptSpecialExcelHandler(scriptType);
            string frameFilePath = Path.Combine(frameOutputDir, $"{SpecialExcelCfg.ErrorCodeFrameScriptFileName}{scriptExcelHandler.Suffix}");
            string businessFilePath = Path.Combine(businessOutputDir, $"{SpecialExcelCfg.ErrorCodeBusinessScriptFileName}{scriptExcelHandler.Suffix}");
            await scriptSpecialHandler.GenerateErrorCodeScript(ErrorCodeDic, frameFilePath, businessFilePath);
        }
    }
}
