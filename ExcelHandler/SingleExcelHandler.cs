namespace DingToolExcelTool.ExcelHandler
{
    using System.Collections.Concurrent;
    using System.IO;
    using System.Text;
    using ClosedXML.Excel;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.ScriptHandler;
    using DingToolExcelTool.Utils;

    internal class SingleExcelHandler : IExcelHandler
    {
        public ConcurrentDictionary<string, SingleExcelHeadInfo> HeadInfoDic { get; private set; }

        public void Init() => HeadInfoDic = new();
        public void Clear() => HeadInfoDic?.Clear();

        public async Task GenerateExcelHeadInfo(string excelInputFile)
        {
            if (HeadInfoDic == null) throw new Exception("[GenerateExcelHeadInfo] HeadInfoDic == null");
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelHeadInfo] 表路径不存在：{excelInputFile}");

            string excelFileName = Path.GetFileNameWithoutExtension(excelInputFile);
            string excelName = excelFileName.Replace(SpecialExcelCfg.SingleExcelPrefix, string.Empty);
            using XLWorkbook wb = new XLWorkbook(excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                SingleExcelHeadInfo headInfo = new();
                HashSet<string> nameSet = new();
                headInfo.ScriptName = sheetCount == 1 ?  $"{excelName}" : $"{excelName}_{sheet.Name}";
                headInfo.Fields = new(10);

                if (HeadInfoDic.ContainsKey(headInfo.ScriptName)) throw new Exception($"[GenerateExcelHeadInfo] 表：{headInfo.ScriptName} 存在同名的表格");

                bool firstColumn = true;
                foreach (IXLColumn column in sheet.ColumnsUsed())
                {
                    StringBuilder typeName = new(column.Cell(1).GetString());
                    bool isTpyeRow = typeName.Length > 0 && typeName[0].Equals('#');

                    if (isTpyeRow) 
                    {
                        typeName.Remove(0, 1);
                        ParseExcelHead(column, headInfo, nameSet, typeName.ToString(), firstColumn);
                    }
                    else ParseExcelValue(column, headInfo);
                    firstColumn = false;
                }

                headInfo.Trim();
                headInfo.Sort();
                HeadInfoDic.TryAdd(headInfo.ScriptName, headInfo);
            }

            await Task.CompletedTask;
        }

        public async Task GenerateProtoMeta(string metaOutputFile, bool isClient) => await Task.CompletedTask;
        public async Task GenerateProtoScript(string metaInputFile, string protoScriptOutputDir, ScriptTypeEn scriptType) => await Task.CompletedTask;
        public async Task GenerateProtoData(string excelInputFile, string protoDataOutputFile, bool isClient, ScriptTypeEn scriptType) => await Task.CompletedTask;

        public async Task GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType)
        {
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelScript] 表路径不存在：{excelInputFile}");
            if (string.IsNullOrEmpty(excelScriptOutputDir))
            {
                LogMessageHandler.AddWarn($"[GenerateExcelScript] 不存在 Excel Script 的输出路径，将不会执行输出操作");
                return;
            }

            IScriptSpecialExcelHandler scriptHandler = ExcelUtil.GetScriptSpecialExcelHandler(scriptType);
            string excelRelativePath = ExcelUtil.GetExcelRelativePath(excelInputFile);
            string excelRelativeDir = Path.GetDirectoryName(excelRelativePath);
            string scriptDir = Path.Combine(excelScriptOutputDir, excelRelativeDir);

            await scriptHandler.GenerateSingleScript(HeadInfoDic, scriptDir, isClient);
        }


        private void ParseExcelHead(IXLColumn column, SingleExcelHeadInfo headInfo, HashSet<string> nameSet,  string typeName, bool firstColumn)
        {
            if (!Enum.TryParse(typeName.ToString().ToLower(), out HeadType headType)) throw new Exception($"[ParseExceHead] 表：{headInfo.ScriptName} 存在未知的表头字段：{typeName}");

            int fieldIdx = 0;
            foreach(IXLCell cell in column.Cells(false))
            {
                int rowIdx = cell.Address.RowNumber;
                if (rowIdx == 1) continue;
                if (ExcelUtil.IsRearMergedCell(cell)) continue;

                SingleExcelFieldInfo fieldInfo;
                var (startRowIdx, endRowIdx) = ExcelUtil.GetCellRowRange(cell);
                if (firstColumn)
                {
                    fieldInfo = new();
                    headInfo.Fields.Add(fieldInfo);

                    fieldInfo.StartRowIdx = startRowIdx;
                    fieldInfo.EndRowIdx = endRowIdx;
                }
                else
                {
                    bool idxOutofBound = fieldIdx >= headInfo.Fields.Count;
                    if (idxOutofBound) continue;//throw new Exception($"[ParseExceHead] 表：{headInfo.ScriptName}。表头信息没有对齐");

                    fieldInfo = headInfo.Fields[fieldIdx];
                    if (fieldInfo.StartRowIdx != startRowIdx || fieldInfo.EndRowIdx != endRowIdx) throw new Exception($"[ParseExceHead] 表：{headInfo.ScriptName}。表头信息没有对齐");
                }
                fieldIdx++;

                string rowContent = cell.GetString().Trim();
                switch (headType)
                {
                    case HeadType.name:
                        if (string.IsNullOrEmpty(rowContent))
                        {
                            LogMessageHandler.AddWarn($"[ParseExceHead] 表：{headInfo.ScriptName} 名字是空的，将不会导出。 Address: {cell.Address}");
                            continue;
                        }

                        if (!nameSet.Add(rowContent)) throw new Exception($"[ParseExceHead] 表：{headInfo.ScriptName}。出现了同名的字段：{rowContent}");

                        fieldInfo.Name = rowContent;
                        break;
                    case HeadType.type:
                        if (string.IsNullOrEmpty(rowContent))
                        {
                            LogMessageHandler.AddWarn($"[ParseExceHead] 表：{headInfo.ScriptName} 类型是空的，将不会导出。 Address: {cell.Address}");
                            continue;
                        }

                        string typeStr = ExcelUtil.ClearKeySymbol(rowContent);
                        if (!ExcelUtil.IsValidType(typeStr)) throw new Exception($"[ParseExceHead] 表：{headInfo.ScriptName} 类型不合法：{typeStr}; 只能是基础数据类型：int, long, double, bool, string, dateTime; 以及预定义的枚举类型; 或者是数组和字典");

                        fieldInfo.Type = typeStr;
                        fieldInfo.LocalizationTxt = ExcelUtil.IsTypeLocalizationTxt(typeStr);
                        fieldInfo.LocalizationImg = ExcelUtil.IsTypeLocalizationImg(typeStr);
                        break;
                    case HeadType.platform:
                        string columnContentLower = rowContent.ToLower();
                        fieldInfo.Platform = columnContentLower switch
                        {
                            "c" => PlatformType.Client,
                            "s" => PlatformType.Server,
                            "cs" => PlatformType.All,
                            _ => PlatformType.Empty,
                        };

                        if (fieldInfo.Platform == PlatformType.Empty)
                        {
                            LogMessageHandler.AddWarn($"[GenerateExcelHeadInfo] 表：{headInfo.ScriptName} Platform 不是合法的字符串，将不会导出。 Address: {cell.Address}");
                            continue;
                        }
                        break;
                    case HeadType.comment:
                        fieldInfo.Comment = rowContent;
                        break;
                }
            }
        }

        private void ParseExcelValue(IXLColumn column, SingleExcelHeadInfo headInfo)
        {
            foreach (IXLCell cell in column.Cells(false))
            {
                int rowIdx = cell.Address.RowNumber;
                if (rowIdx == 1) continue;
                if (ExcelUtil.IsRearMergedCell(cell)) continue;

                int fieldIdx = headInfo.GetFieldIdx(rowIdx);
                if (fieldIdx == -1) throw new Exception($"[GenerateProtoData] 表：{headInfo.ScriptName} 存在字段没有和类型关联上. Address: {cell.Address}");
                if (!string.IsNullOrEmpty(headInfo.Fields[fieldIdx].Value)) throw new Exception($"[GenerateProtoData] 表：{headInfo.ScriptName} 是一个单例表，只能有一列的数据");

                string columnContent = cell.GetString().Trim();
                headInfo.Fields[fieldIdx].Value = columnContent;
            }
        }
    }
}
