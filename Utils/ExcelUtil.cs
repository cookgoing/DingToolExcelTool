namespace DingToolExcelTool.Utils
{
    using System.IO;
    using System.Linq;
    using System.Diagnostics;
    using ClosedXML.Excel;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.ExcelHandler;
    using DingToolExcelTool.ScriptHandler;

    internal static class ExcelUtil
    {
        public static KeyType GetKeyType(string typeStr)
        {
            if (typeStr.StartsWith(GeneralCfg.UnionKeySymbol)) return KeyType.Union;
            if (typeStr.StartsWith(GeneralCfg.IndependentKeySymbol)) return KeyType.Independent;

            return KeyType.No;
        }
        public static string ClearKeySymbol(string typeStr)
        {
            KeyType keyType = GetKeyType(typeStr);
            if (keyType == KeyType.Union) return typeStr.Replace(GeneralCfg.UnionKeySymbol, string.Empty);
            if (keyType == KeyType.Independent) return typeStr.Replace(GeneralCfg.IndependentKeySymbol, string.Empty);

            return typeStr;
        }


        public static bool IsValidType(string typeStr) => IsTypeLocalizationTxt(typeStr) || IsTypeLocalizationImg(typeStr) || IsBaseType(typeStr) || IsArrType(typeStr) || IsMapType(typeStr) || IsEnumType(typeStr);
        public static bool IsTypeLocalizationTxt(string typeStr) => typeStr == GeneralCfg.LocalizationTxtSymbol;
        public static bool IsTypeLocalizationImg(string typeStr) => typeStr == GeneralCfg.LocalizationImgSymbol;
        public static bool IsBaseType(string typeStr) => GeneralCfg.ExcelBaseType.Contains(typeStr);
        public static bool IsEnumType(string typeStr) => ExcelManager.Instance.EnumExcelHandler.EnumDic.ContainsKey(typeStr);
        public static bool IsArrType(string typeStr)
        {
            if (!typeStr.EndsWith("[]")) return false;

            string elementType = typeStr.Substring(0, typeStr.Length - 2);
            return IsBaseType(elementType) || IsEnumType(elementType);
        }
        public static bool IsMapType(string typeStr)
        {
            if (!typeStr.StartsWith("map<") || !typeStr.EndsWith('>')) return false;

            string innerTypes = typeStr.Substring(4, typeStr.Length - 5);
            string[] keyValue = innerTypes.Split(',');

            if (keyValue.Length != 2) return false;

            return (IsBaseType(keyValue[0]) || IsEnumType(keyValue[0])) && (IsBaseType(keyValue[1]) || IsEnumType(keyValue[1]));
        }


        public static string? ToProtoType(string typeStr)
        {
            if (IsTypeLocalizationTxt(typeStr) || IsTypeLocalizationImg(typeStr)) return "string";
            if (IsBaseType(typeStr)) return GeneralCfg.BaseType2ProtoMap[typeStr];
            if (IsArrType(typeStr))
            {
                string elementType = typeStr.Substring(0, typeStr.Length - 2);
                if (GeneralCfg.BaseType2ProtoMap.TryGetValue(elementType, out string protoType)) return $"repeated {protoType}";
                else return $"repeated {elementType}";
            }
            if (IsMapType(typeStr))
            {
                string innerTypes = typeStr.Substring(4, typeStr.Length - 5);
                string[] keyValue = innerTypes.Split(',');
                string kType = keyValue[0], vType = keyValue[1];

                if (GeneralCfg.BaseType2ProtoMap.TryGetValue(kType, out string? value)) kType = value;
                if (GeneralCfg.BaseType2ProtoMap.TryGetValue(vType, out value)) vType = value;

                return $"map<{kType},{vType}>";
            }
            if (IsEnumType(typeStr))
            {
                return $"{typeStr}";
            }

            return null;
        }


        public static bool IsRearMergedCell(IXLCell cell)
        {
            if (cell == null) return false;

            bool isMerged = cell.IsMerged();
            if (!isMerged) return false;

            IXLWorksheet sheet = cell.Worksheet;
            IXLRange mergedRange = sheet.MergedRanges.FirstOrDefault(range => range.Contains(cell));

            return mergedRange != null && !cell.Address.Equals(mergedRange.FirstCell().Address);
        }

        public static (int startColumnIdx, int endColumnIdx) GetCellColumnRange(IXLCell cell)
        {
            bool isMerged = cell.IsMerged();
            if (!isMerged) return (cell.Address.ColumnNumber, cell.Address.ColumnNumber);

            IXLWorksheet sheet = cell.Worksheet;
            IXLRange mergedRange = sheet.MergedRanges.FirstOrDefault(range => range.Contains(cell));
            if (mergedRange == null) return (cell.Address.ColumnNumber, cell.Address.ColumnNumber);

            return (mergedRange.FirstCell().Address.ColumnNumber, mergedRange.LastCell().Address.ColumnNumber);
        }

        public static (int startColumnIdx, int endColumnIdx) GetCellRowRange(IXLCell cell)
        {
            bool isMerged = cell.IsMerged();
            if (!isMerged) return (cell.Address.RowNumber, cell.Address.RowNumber);

            IXLWorksheet sheet = cell.Worksheet;
            IXLRange mergedRange = sheet.MergedRanges.FirstOrDefault(range => range.Contains(cell));
            if (mergedRange == null) return (cell.Address.RowNumber, cell.Address.RowNumber);

            return (mergedRange.FirstCell().Address.RowNumber, mergedRange.LastCell().Address.RowNumber);
        }

        public static string? GetExcelRelativePath(string excelFilePath)
        {
            string excelRootDirPath = ExcelManager.Instance.Data?.ExcelInputRootDir;
            if (excelRootDirPath == null || !excelFilePath.StartsWith(excelRootDirPath)) return null;

            return Path.GetRelativePath(excelRootDirPath, excelFilePath);
        }


        public static void GenerateProtoScript(string arguments) => ExcuteProgramFile(GeneralCfg.ProtocPath, arguments);

        public static void ExcuteProgramFile(string batchFilePath, string arguments)
        {
            if (!File.Exists(batchFilePath)) throw new FileNotFoundException($"[ExcuteBatchFile] 没有这个应用程序: {batchFilePath}", batchFilePath);

            ProcessStartInfo startInfo = new()
            {
                FileName = batchFilePath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using Process process = new() { StartInfo = startInfo };

            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (process.ExitCode != 0) throw new Exception($"【ExcuteBatchFile】执行失败：{error}");
        }


        public static IScriptExcelHandler GetScriptExcelHandler(ScriptTypeEn scriptType)
        {
            return scriptType switch
            {
                ScriptTypeEn.CSharp => CSharpExcelHandler.Instance,
                _ => throw new Exception($"[GetScriptHandler]. 未知的脚本类型：{scriptType}")
            };
        }

        public static IScriptSpecialExcelHandler GetScriptSpecialExcelHandler(ScriptTypeEn scriptType)
        {
            return scriptType switch
            {
                ScriptTypeEn.CSharp => CSharpSpecialExcelHandler.Instance,
                _ => throw new Exception($"[GetScriptHandler]. 未知的脚本类型：{scriptType}")
            };
        }


        public static void ClearDirectory(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath))
                {
                    foreach (string file in Directory.GetFiles(folderPath)) File.Delete(file);

                    foreach (string dir in Directory.GetDirectories(folderPath))
                    {
                        ClearDirectory(dir);
                        Directory.Delete(dir);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"清空文件夹时发生错误: {ex.Message}");
            }
        }

        public static string ParsePath(string path) => path.Replace("%BaseDir%", AppDomain.CurrentDomain.BaseDirectory);
    }
}
