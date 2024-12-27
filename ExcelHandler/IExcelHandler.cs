namespace DingToolExcelTool.ExcelHandler
{
    using DingToolExcelTool.Data;

    internal interface IExcelHandler
    {
        void Init();
        void Clear();

        void GenerateExcelHeadInfo(string excelInputFile);
        void GenerateProtoMeta(string metaOutputFile, bool isClient);
        void GenerateProtoScript(string metaInputFile, string protoScriptOutputDir, ScriptTypeEn scriptType);
        void GenerateProtoData(string excelInputFile, string protoDataOutputDir, bool isClient, ScriptTypeEn scriptType);
        void GenerateExcelScript(string excelInputFile, string excelScriptOutputDir, bool isClient, ScriptTypeEn scriptType);
    }
}
