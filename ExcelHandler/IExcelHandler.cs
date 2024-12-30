namespace DingToolExcelTool.ExcelHandler
{
    using DingToolExcelTool.Data;

    internal interface IExcelHandler
    {
        void Init();
        void Clear();

        Task GenerateExcelHeadInfo(string excelInputFile);
        Task GenerateProtoMeta(string metaOutputFile, bool isClient);
        Task GenerateProtoScript(string metaInputFile, string protoScriptOutputDir, ScriptTypeEn scriptType);
        Task GenerateProtoData(string excelInputFile, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType);
        Task GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType);
    }
}
