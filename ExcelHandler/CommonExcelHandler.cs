namespace DingToolExcelTool.ExcelHandler
{
    using System.IO;
    using System.Text;
    using System.Collections.Concurrent;
    using ClosedXML.Excel;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Utils;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.ScriptHandler;
    
    internal class CommonExcelHandler : IExcelHandler
    {
        public ConcurrentDictionary<string, ExcelHeadInfo> HeadInfoDic { get; private set; }

        public virtual void Init() => HeadInfoDic = new();

        public virtual void Clear() => HeadInfoDic?.Clear();

        public virtual void GenerateExcelHeadInfo(string excelInputFile)
        {
            if (HeadInfoDic == null) throw new Exception("[GenerateExcelHeadInfo] HeadInfoDic == null");
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelHeadInfo] 表路径不存在：{excelInputFile}");

            string excelFileName = Path.GetFileNameWithoutExtension(excelInputFile);
            using XLWorkbook wb = new XLWorkbook(excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                ExcelHeadInfo headInfo = new();
                HashSet<string> nameSet = new();
                headInfo.MessageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
                headInfo.Fields = new (10);
                headInfo.UnionKey = new();
                headInfo.IndependentKey = new();

                if (HeadInfoDic.ContainsKey(headInfo.MessageName)) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName} 存在同名的表格");

                bool firstRow = true;
                foreach (IXLRow row in sheet.RowsUsed())
                {
                    StringBuilder typeName = new(row.Cell(1).GetString());
                    if (typeName.Length == 0 || !typeName[0].Equals('#')) continue;
                    
                    typeName.Remove(0, 1);
                    if (!Enum.TryParse(typeName.ToString().ToLower(), out HeadType headType)) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName} 存在未知的表头字段：{typeName}");

                    int fieldIdx = 0;
                    foreach (IXLCell cell in row.Cells(false))
                    {
                        int columnIdx = cell.Address.ColumnNumber;
                        if (columnIdx == 1) continue;
                        if (ExcelUtil.IsRearMergedCell(cell)) continue;

                        string columnContent = cell.GetString().Trim();
                        ExcelFieldInfo fieldInfo;
                        var (startColumnIdx, endColumnIdx) = ExcelUtil.GetCellColumnRange(cell);
                        if (firstRow)
                        {
                            fieldInfo = new();
                            headInfo.Fields.Add(fieldInfo);

                            fieldInfo.StartColumnIdx = startColumnIdx;
                            fieldInfo.EndColumnIdx = endColumnIdx;
                        }
                        else
                        {
                            bool idxOutofBound = fieldIdx >= headInfo.Fields.Count;
                            if (idxOutofBound) continue;

                            fieldInfo = headInfo.Fields[fieldIdx];
                            if (fieldInfo.StartColumnIdx != startColumnIdx || fieldInfo.EndColumnIdx != endColumnIdx) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName}。表头信息没有对齐: {cell.Address}; fieldIdx: {fieldIdx}; cell: ({startColumnIdx}, {endColumnIdx}); field: ({fieldInfo.StartColumnIdx},{fieldInfo.EndColumnIdx})");
                        }
                        fieldIdx++;

                        switch (headType)
                        {
                            case HeadType.name:
                                if (string.IsNullOrEmpty(columnContent))
                                {
                                    LogMessageHandler.AddWarn($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName} 名字是空的，将不会导出。 Address: {cell.Address}");
                                    continue;
                                }

                                if (!nameSet.Add(columnContent)) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName}。出现了同名的字段：{columnContent}");

                                fieldInfo.Name = columnContent;
                                break;
                            case HeadType.type:
                                if (string.IsNullOrEmpty(columnContent))
                                {
                                    LogMessageHandler.AddWarn($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName} 类型是空的，将不会导出。 Address: {cell.Address}");
                                    continue;
                                }

                                KeyType keyType = ExcelUtil.GetKeyType(columnContent);
                                switch (keyType)
                                {
                                    case KeyType.Independent: headInfo.IndependentKey.Add(fieldInfo); break;
                                    case KeyType.Union: headInfo.UnionKey.Add(fieldInfo); break;
                                }

                                string typeStr = ExcelUtil.ClearKeySymbol(columnContent);
                                if (!ExcelUtil.IsValidType(typeStr)) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName}, address: {cell.Address} 类型不合法：{typeStr}; 只能是基础数据类型：int, long, double, bool, string, dateTime; 以及预定义的枚举类型; 或者是数组和字典");

                                fieldInfo.Type = typeStr;
                                fieldInfo.LocalizationTxt = ExcelUtil.IsTypeLocalizationTxt(typeStr);
                                fieldInfo.LocalizationImg = ExcelUtil.IsTypeLocalizationImg(typeStr);
                                break;
                            case HeadType.platform:
                                string columnContentLower = columnContent.ToLower();
                                fieldInfo.Platform = columnContentLower switch
                                {
                                    "c" => PlatformType.Client,
                                    "s" => PlatformType.Server,
                                    "cs" => PlatformType.All,
                                    _ => PlatformType.Empty,
                                };

                                if (fieldInfo.Platform == PlatformType.Empty)
                                {
                                    LogMessageHandler.AddWarn($"[GenerateExcelHeadInfo] 表：{headInfo.MessageName} Platform 没有指定平台，将不会导出。 Address: {cell.Address}");
                                    continue;
                                }
                                break;
                            case HeadType.comment:
                                fieldInfo.Comment = columnContent;
                                break;
                        }
                    }

                    firstRow = false;
                }

                headInfo.Trim();
                headInfo.Sort();
                HeadInfoDic.TryAdd(headInfo.MessageName, headInfo);
            }
        }

        public virtual void GenerateProtoMeta(string metaOutputFile, bool isClient)
        {
            LogMessageHandler.AddInfo($"[CommonExcelHandler][GenerateProtoMeta]: {metaOutputFile}, isClient: {isClient}");
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

            foreach (ExcelHeadInfo headInfo in HeadInfoDic.Values)
            {
                metaWriter.AppendLine($"message {headInfo.MessageName}").AppendLine("{");
                int messageFieldIdx = 0;
                foreach (ExcelFieldInfo fieldInfo in headInfo.Fields)
                {
                    if ((platform & fieldInfo.Platform) == 0) continue;

                    metaWriter.Append($"\t{ExcelUtil.ToProtoType(fieldInfo.Type)} {fieldInfo.Name} = {++messageFieldIdx};");
                    if (string.IsNullOrEmpty(fieldInfo.Comment)) metaWriter.AppendLine();
                    else metaWriter.AppendLine($"//{fieldInfo.Comment}");
                }
                metaWriter.AppendLine("}").AppendLine();
            }

            metaSW.Write(metaWriter);
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

            foreach (ExcelHeadInfo headInfo in HeadInfoDic.Values)
            {
                listWriter.AppendLine(@$"message {headInfo.MessageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}
{{
    repeated {headInfo.MessageName} {CommonExcelCfg.ProtoMetaListFieldName} = 1;
}}").AppendLine();
            }

            listSW.Write(listWriter);
            listSW.Close();
            listWriter.Clear();
        }

        public virtual void GenerateProtoScript(string metaInputFile, string protoScriptOutputDir, ScriptTypeEn scriptType)
        {
            LogMessageHandler.AddInfo($"[CommonExcelHandler][GenerateProtoScript]: output: {protoScriptOutputDir}, scriptType: {scriptType}");

            if (!File.Exists(metaInputFile)) throw new Exception($"[GenerateProtoScript] proto meta 文件路径不存在：{metaInputFile}");
            if (string.IsNullOrEmpty(protoScriptOutputDir))
            {
                LogMessageHandler.AddWarn($"[GenerateProtoScript] 不存在 proto script 的输出路径，将不会执行输出操作");
                return;
            }

            IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
            scriptHandler.GenerateProtoScript(metaInputFile, protoScriptOutputDir);

            string dirPath = Path.GetDirectoryName(metaInputFile);
            string metaFileName = Path.GetFileName(metaInputFile);
            string metaName = metaFileName.Replace(GeneralCfg.ProtoMetaFileSuffix, string.Empty);
            string listFilePath = Path.Combine(dirPath, $"{metaName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}{GeneralCfg.ProtoMetaFileSuffix}");
            if (!Path.Exists(listFilePath)) return;

            scriptHandler.GenerateProtoScript(listFilePath, protoScriptOutputDir);
        }

        public virtual void GenerateProtoData(string excelInputFile, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType)
        {
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateProtoData] 表路径不存在：{excelInputFile}");
            if (string.IsNullOrEmpty(protoDataOutputDir))
            {
                LogMessageHandler.AddWarn($"[GenerateProtoData] 不存在 proto data 的输出路径，将不会执行输出操作");
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

                foreach (IXLRow row in sheet.RowsUsed())
                {
                    string typeName = row.Cell(1).GetString();
                    bool isTpyeRow = typeName.StartsWith('#');
                    if (isTpyeRow) continue;

                    foreach (IXLCell cell in row.Cells(false))
                    {
                        int columnIdx = cell.Address.ColumnNumber;
                        if (columnIdx == 1) continue;
                        if (ExcelUtil.IsRearMergedCell(cell)) continue;

                        int fieldIdx = headInfo.GetFieldIdx(columnIdx);
                        if (fieldIdx == -1) continue;

                        ExcelFieldInfo fieldInfo = headInfo.Fields[fieldIdx];
                        if ((platform & fieldInfo.Platform) == 0) continue;

                        string columnContent = cell.GetString().Trim();

                        //LogMessageHandler.AddWarn($"[test]. messageName: {messageName}; Address: {cell.Address}; type: {fieldInfo.Type}; columnContent: {columnContent}; platform: {platform}; fieldPlatform: {fieldInfo.Platform}; isClient: {isClient}");
                        if (ExcelUtil.IsTypeLocalizationTxt(fieldInfo.Type)
                            || ExcelUtil.IsTypeLocalizationImg(fieldInfo.Type)
                            || ExcelUtil.IsBaseType(fieldInfo.Type))
                        {
                            scriptHandler.SetScriptValue(messageName, fieldInfo.Name, fieldInfo.Type, columnContent);
                        }
                        else if (ExcelUtil.IsArrType(fieldInfo.Type))
                        {
                            scriptHandler.AddScriptList(messageName, fieldInfo.Name, fieldInfo.Type, columnContent);
                        }
                        else if (ExcelUtil.IsMapType(fieldInfo.Type))
                        {
                            int relativeColumnIdx = columnIdx - fieldInfo.StartColumnIdx;
                            bool isKey = relativeColumnIdx % 2 == 0;
                            if (!isKey) continue;

                            IXLCell nextCell = row.Cell(++columnIdx);
                            fieldIdx = headInfo.GetFieldIdx(columnIdx);
                            if (nextCell == null || fieldIdx == -1 || !ExcelUtil.IsMapType(headInfo.Fields[fieldIdx].Type)) throw new Exception($"[GenerateProtoData] 表：{messageName} 没有这个字段：{fieldInfo.Name} || 格式不合法，这里应该是一个字典值。 Address: {nextCell?.Address}");

                            string keyData = columnContent;
                            string valueData = nextCell.GetString().Trim();
                            scriptHandler.AddScriptMap(messageName, fieldInfo.Name, fieldInfo.Type, keyData, valueData);
                        }
                    }

                    scriptHandler.AddListScriptObj($"{messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", messageName);
                }

                string protoDataOutputFile = Path.Combine(protoDataOutputDir, $"{messageName}{GeneralCfg.ProtoDataFileSuffix}");
                scriptHandler.SerializeObjInProto($"{messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}", protoDataOutputFile);
            }
        }

        public virtual void GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType)
        {
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelScript] 表路径不存在：{excelInputFile}");
            if (string.IsNullOrEmpty(excelScriptOutputDir))
            {
                LogMessageHandler.AddWarn($"[GenerateExcelScript] 不存在 Excel Script 的输出路径，将不会执行输出操作");
                return;
            }

            IScriptExcelHandler scriptHandler = ExcelUtil.GetScriptExcelHandler(scriptType);
            string excelRelativePath = ExcelUtil.GetExcelRelativePath(excelInputFile);
            string excelFileName = Path.GetFileNameWithoutExtension(excelRelativePath);
            string excelRelativeDir = Path.GetDirectoryName(excelRelativePath);
            using XLWorkbook wb = new(excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                string messageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
                if (!HeadInfoDic.TryGetValue(messageName, out ExcelHeadInfo headInfo)) throw new Exception($"[GenerateProtoData] 表: {messageName} 没有 headInfo");

                string outputFilePath = Path.Combine(excelScriptOutputDir, excelRelativeDir??"", $"{messageName}{GeneralCfg.ExcelScriptFileSuffix(scriptType)}");
                scriptHandler.GenerateExcelScript(headInfo, outputFilePath, isClient);
            }
        }
    }
}
