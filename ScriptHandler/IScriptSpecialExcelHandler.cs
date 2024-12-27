namespace DingToolExcelTool.ScriptHandler
{
    using DingToolExcelTool.Data;

    internal interface IScriptSpecialExcelHandler
    {
        void GenerateErrorCodeScript(Dictionary<string, ErrorCodeScriptInfo> errorCodeHeadDic, string frameOutputFile, string businessOutputFile);

        void GenerateSingleScript(Dictionary<string, SingleExcelHeadInfo> singleHeadDic, string outputDir, bool isClient);
    }
}
