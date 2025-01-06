namespace DingToolExcelTool.ScriptHandler
{
    using DingToolExcelTool.Data;
    using System.Collections.Concurrent;

    internal interface IScriptSpecialExcelHandler
    {
        Task GenerateErrorCodeScript(ConcurrentDictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile);

        Task GenerateSingleScript(SingleExcelHeadInfo singleHeadInfo, string outputFile, bool isClient);
    }
}
