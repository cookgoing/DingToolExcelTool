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

                if (!string.IsNullOrEmpty(scriptInfo.SheetName)) ErrorCodeDic.TryAdd(scriptInfo.SheetName, scriptInfo);
            }
        }

        public override async Task GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType) 
        {
            await base.GenerateExcelScript(excelInputFile, excelScriptOutputDir, isClient, scriptType);

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
