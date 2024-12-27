namespace DingToolExcelTool.ScriptHandler
{
    using System.IO;
    using System.Text;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using System.Reflection;

    internal class CSharpSpecialExcelHandler : Singleton<CSharpSpecialExcelHandler>, IScriptSpecialExcelHandler
    {
        public void GenerateErrorCodeScript(Dictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile)
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
                    if (isFrame) frameFieldSB.AppendLine($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};//{fieldInfo.Comment}");
                    else businessFieldSB.AppendLine($"\t\tpublic const int {fieldInfo.CodeStr} = {fieldInfo.Code};//{fieldInfo.Comment}");
                }
            }

            WriteErrorCodeScript(frameOutputFile, frameFieldSB.ToString(), SpecialExcelCfg.ErrorCodeFramePackageName);
            WriteErrorCodeScript(businessOutputFile, businessFieldSB.ToString(), SpecialExcelCfg.ErrorCodeBusinessPackageName);
        }

        public void GenerateSingleScript(Dictionary<string, SingleExcelHeadInfo> singleHeadDic, string outputDir, bool isClient)
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

                    fieldSB.AppendLine($"\t\tpublic readonly static {CSharpExcelHandler.Instance.ExcelType2ScriptTypeStr(fieldInfo.Type)} {fieldInfo.Name} = {fieldInfo.Value};//{fieldInfo.Comment}");
                }

                string filePath = Path.Combine(outputDir, singleHeadInfo.ScriptName, $"{GeneralCfg.ExcelScriptFileSuffix(ScriptTypeEn.CSharp)}");
                WriteSingleExcelScript(singleHeadInfo.ScriptName, filePath, fieldSB.ToString());
            }
        }

        private void WriteErrorCodeScript(string outputPath, string fieldStr, string packageName)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
using System.IO
using System.Collections.Generic;
using Google.Protobuf;
using DingFrame.Module.AssetLoader;

namespace {packageName};
public sealed partial class {SpecialExcelCfg.ErrorCodeScriptName}
{{
    {fieldStr}
}}
");

            sw.Write(scriptSB.ToString());
            sw.Flush();
        }

        private void WriteSingleExcelScript(string scriptName, string outputPath, string fieldStr)
        {
            using StreamWriter sw = new StreamWriter(outputPath);
            StringBuilder scriptSB = new();

            scriptSB.AppendLine(@$"
using System.IO
using System.Collections.Generic;
using Google.Protobuf;
using DingFrame.Module.AssetLoader;

namespace {GeneralCfg.ProtoMetaPackageName};
public static class {scriptName}
{{
    {fieldStr}
}}
");

            sw.Write(scriptSB.ToString());
            sw.Flush();
        }
    }
}
