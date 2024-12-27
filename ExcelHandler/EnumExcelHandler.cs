namespace DingToolExcelTool.ExcelHandler
{
    using System.IO;
    using System.Text;
    using ClosedXML.Excel;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Utils;
    
    internal class EnumExcelHandler : CommonExcelHandler
    {
        public Dictionary<string, EnumInfo> EnumDic { get; private set; }//<EnumName, info>

        public override void Init()
        {
            base.Init();
            EnumDic = new();
        }

        public override void Clear()
        {
            base.Clear();
            EnumDic?.Clear();
        }

        public override void GenerateExcelHeadInfo(string excelInputFile)
        {
            if (EnumDic == null) throw new Exception("[GenerateExcelHeadInfo] EnunDic == null");
            if (!File.Exists(excelInputFile)) throw new Exception($"[GenerateExcelHeadInfo] 表路径不存在：{excelInputFile}");

            base.GenerateExcelHeadInfo(excelInputFile);

            string excelFileName = Path.GetFileNameWithoutExtension(excelInputFile);
            using XLWorkbook wb = new (excelInputFile);
            int sheetCount = wb.Worksheets.Count;

            foreach (IXLWorksheet sheet in wb.Worksheets)
            {
                string headMessageName = sheetCount == 1 ? $"{excelFileName}" : $"{excelFileName}_{sheet.Name}";
                if (!HeadInfoDic.TryGetValue(headMessageName, out ExcelHeadInfo headInfo)) throw new Exception($"没有[{headMessageName}] 的表头信息");

                List<string> enumFieldNames = new(SpecialExcelCfg.EnumFixedField.Keys);
                headInfo.Fields.ForEach(fieldInfo =>
                {
                    if (SpecialExcelCfg.EnumFixedField.TryGetValue(fieldInfo.Name, out string typeName))
                    {
                        if (typeName != fieldInfo.Type) throw new Exception($"[GenerateExcelHeadInfo]. 枚举表中有一个字段的类型不一致：{fieldInfo.Name} - {fieldInfo.Type}; 类型应该是：{typeName}");

                        enumFieldNames.Remove(fieldInfo.Name);
                    }
                });

                StringBuilder sb = new();
                foreach (string field in enumFieldNames) sb.Append(field).Append(',');
                if (enumFieldNames.Count > 0)
                {
                    sb.Remove(sb.Length - 1, 1);
                    throw new Exception($"[GenerateExcelHeadInfo] 枚举表有缺失的字段: {sb}");
                }

                int columnCount = sheet.ColumnCount();
                EnumInfo enumInfo = new();
                HashSet<string> nameSet = new();
                foreach (IXLRow row in sheet.RowsUsed())
                {
                    int columnIdx = 1;
                    string typeName = row.Cell(columnIdx).GetString();
                    bool isTpyeRow = typeName.StartsWith('#');
                    if (!isTpyeRow) continue;

                    foreach (IXLCell cell in row.CellsUsed())
                    {
                        if (ExcelUtil.IsRearMergedCell(cell)) continue;

                        columnIdx = cell.Address.ColumnNumber;
                        int fieldIdx = headInfo.GetFieldIdx(columnIdx);
                        if (fieldIdx == -1) throw new Exception($"[GenerateExcelHeadInfo] 表：{headMessageName} 存在字段没有和类型关联上. Address: {cell.Address}");

                        ExcelFieldInfo fieldInfo = headInfo.Fields[fieldIdx];
                        string columnContent = cell.GetString().Trim();
                        switch (fieldInfo.Name.ToLower())
                        {
                            case "name":
                                if (!nameSet.Add(columnContent)) throw new Exception($"[GenerateExcelHeadInfo]. 枚举表中出现了相同名字的枚举类型：{columnContent}");
                                enumInfo.Name = columnContent; 
                                break;
                            case "field":
                                string[] fieldStrs = columnContent.Split(SpecialExcelCfg.EnumFieldSplitSymbol);
                                enumInfo.Fields = new EnumFieldInfo[fieldStrs.Length - 1];
                                for (int i = 0; i < fieldStrs.Length; ++i) enumInfo.Fields[i].Name = fieldStrs[i];
                                break;
                            case "value":
                                string[] valueStrs = columnContent.Split(SpecialExcelCfg.EnumFieldSplitSymbol);
                                if (enumInfo.Fields?.Length != valueStrs.Length) throw new Exception($"[GenerateExcelHeadInfo] 枚举[{enumInfo.Name}] 格式不合规：字段名字和字段的数值数量没有统一");
                                for (int i = 0; i < valueStrs.Length; ++i)
                                {
                                    if (!int.TryParse(valueStrs[i], out int value)) throw new Exception($"[GenerateExcelHeadInfo] 枚举的value字段存在不能解析成int的字段：{valueStrs[i]}");
                                    enumInfo.Fields[i].Value = value;
                                }
                                break;
                            case "platform":
                                string columnContentLower = columnContent.ToLower();
                                enumInfo.Platform = columnContentLower switch
                                {
                                    "c" => PlatformType.Client,
                                    "s" => PlatformType.Server,
                                    "cs" => PlatformType.All,
                                    _ => PlatformType.Empty,
                                };
                                break;
                            case "comment":
                                enumInfo.Comment = columnContent;
                                break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(enumInfo.Name)) EnumDic.Add(enumInfo.Name, enumInfo);
            }
        }

        public override void GenerateProtoMeta(string metaOutputFile, bool isClient)
        {
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
package {GeneralCfg.ProtoMetaPackageName};").AppendLine();

            foreach (EnumInfo enumInfo in EnumDic.Values)
            {
                if ((platform & enumInfo.Platform) == 0) continue;

                metaWriter.AppendLine($"//{enumInfo.Comment}");
                metaWriter.AppendLine($"message {enumInfo.Name}").AppendLine("{");
                foreach (EnumFieldInfo fieldInfo in enumInfo.Fields)
                {
                    metaWriter.AppendLine($"\t{fieldInfo.Name} = {fieldInfo.Value};");
                }
                metaWriter.AppendLine("}").AppendLine();
            }

            metaSW.Write(metaWriter);
            metaSW.Close();
            metaWriter.Clear();
        }

        public override void GenerateProtoData(string excelInputFile, string protoDataOutputFile, bool isClient, ScriptTypeEn scriptType) { }
        public override void GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType) { }
    }
}
