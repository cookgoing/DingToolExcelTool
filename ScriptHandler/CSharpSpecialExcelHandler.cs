namespace DingToolExcelTool.ScriptHandler
{
    using System.IO;
    using System.Text;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using System.Collections.Concurrent;

    internal class CSharpSpecialExcelHandler : Singleton<CSharpSpecialExcelHandler>, IScriptSpecialExcelHandler
    {
        public async Task GenerateErrorCodeScript(ConcurrentDictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile)
        {
            if (string.IsNullOrEmpty(frameOutputFile) || string.IsNullOrEmpty(businessOutputFile)) throw new Exception("[GenerateErrorCodeScript]. ErrorCode 需要有两个输出路径");
            if (errorCodeHeadDic == null) throw new Exception("[GenerateErrorCodeScript]. Errorcode 表，没有头部信息");

            string frameOutputDir = Path.GetDirectoryName(frameOutputFile);
            string businessOutputDir = Path.GetDirectoryName(businessOutputFile);
            if (!Directory.Exists(frameOutputDir)) Directory.CreateDirectory(frameOutputDir);
            if (!Directory.Exists(businessOutputDir)) Directory.CreateDirectory(businessOutputDir);

            StringBuilder frameFieldSB = new();
            StringBuilder businessFieldSB = new();
            foreach (ErrorCodeScriptInfo errorCodeInfo in errorCodeHeadDic.Values)
            {
                bool isFrame = errorCodeInfo.SheetName == SpecialExcelCfg.ErrorCodeFrameSheetName;
                foreach (ErrorCodeScriptFieldInfo fieldInfo in errorCodeInfo.Fields)
                {
                    if (isFrame) frameFieldSB.Append($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                    else businessFieldSB.Append($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment);
                }
            }

            await WriteErrorCodeScript(frameOutputFile, frameFieldSB.ToString(), SpecialExcelCfg.ErrorCodeFramePackageName);
            await WriteErrorCodeScript(businessOutputFile, businessFieldSB.ToString(), SpecialExcelCfg.ErrorCodeBusinessPackageName);
        }

        public async Task GenerateSingleScript(ConcurrentDictionary<string, SingleExcelHeadInfo> singleHeadDic, string outputDir, bool isClient)
        {
            if (string.IsNullOrEmpty(outputDir)) throw new Exception("[GenerateSingleScript]. 没有输出路径");

            PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            foreach (SingleExcelHeadInfo singleHeadInfo in singleHeadDic.Values)
            {
                StringBuilder fieldSB = new();
                foreach (SingleExcelFieldInfo fieldInfo in singleHeadInfo.Fields)
                {
                    if ((platform & fieldInfo.Platform) == 0) continue;

                    string filedValue = CSharpExcelHandler.Instance.ExcelType2ScriptType(fieldInfo.Type, fieldInfo.Value).ToString();
                    if (fieldInfo.Type == "string") filedValue = $"\"{filedValue}\"";
                    else if (fieldInfo.Type == "bool") filedValue = filedValue?.ToLower();
                    fieldSB.Append($"\t\tpublic readonly static {CSharpExcelHandler.Instance.ExcelType2ScriptTypeStr(fieldInfo.Type)} {fieldInfo.Name} = {filedValue};").AppendLine(string.IsNullOrEmpty(fieldInfo.Comment) ? null : "//" + fieldInfo.Comment) ;
                }

                string filePath = Path.Combine(outputDir, $"{singleHeadInfo.ScriptName}{GeneralCfg.ExcelScriptFileSuffix(ScriptTypeEn.CSharp)}");
                await WriteSingleExcelScript(singleHeadInfo.ScriptName, filePath, fieldSB.ToString());
            }
        }

        private async Task WriteErrorCodeScript(string outputPath, string fieldStr, string packageName)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
namespace {packageName}
{{
    public sealed partial class {SpecialExcelCfg.ErrorCodeScriptName}
    {{
{fieldStr}
    }}
}}
");

            await sw.WriteAsync(scriptSB.ToString());
            sw.Flush();
        }

        private async Task WriteSingleExcelScript(string scriptName, string outputPath, string fieldStr)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
using System.Collections.Generic;

namespace {GeneralCfg.ProtoMetaPackageName}
{{
    public static class {scriptName}
    {{
{fieldStr}
    }}
}}
");

            await sw.WriteAsync(scriptSB.ToString());
            sw.Flush();
        }
    }
}
