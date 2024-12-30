namespace DingToolExcelTool.ScriptHandler
{
    using DingToolExcelTool.Data;

    internal interface IScriptExcelHandler
    {
        public Dictionary<string, string> BaseType2ScriptMap { get; }
        public string Suffix { get; }

        void DynamicCompile(string[] csCodes);

        string ExcelType2ScriptTypeStr(string typeStr);

        void GenerateProtoScript(string metaInputFile, string protoScriptOutputDir);
        void GenerateProtoScriptBatchly(string metaInputDir, string protoScriptOutputDir);

        void SetScriptValue(string scriptName, string fieldName, string typeStr, string valueStr);
        void AddScriptList(string scriptName, string fieldName, string typeStr, string valueStr);
        void AddScriptMap(string scriptName, string fieldName, string typeStr, string keyData, string valueData);
        void AddListScriptObj(string scriptName, string itemScriptName);
        void SerializeObjInProto(string scriptName, string outputFilePath);

        Task GenerateExcelScript(ExcelHeadInfo headInfo, string excelScriptOutputFile, bool isClient);
    }
}
