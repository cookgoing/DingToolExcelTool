namespace DingToolExcelTool.ScriptHandler
{
    using DingToolExcelTool.Data;
    using System.Collections.Concurrent;

    internal interface IScriptSpecialExcelHandler
    {
        void GenerateErrorCodeScript(ConcurrentDictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile);

        void GenerateSingleScript(ConcurrentDictionary<string, SingleExcelHeadInfo> singleHeadDic, string outputDir, bool isClient);
    }
}
