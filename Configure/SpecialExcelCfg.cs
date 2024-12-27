namespace DingToolExcelTool.Configure
{
    internal static class SpecialExcelCfg
    {
        public const string SingleExcelPrefix = "[Single]";
        public const string EnumExcelName = "Enum";
        public const string ErrorCodeExcelName = "ErrorCode";

        public const string EnumProtoMetaFileName = "EnumExcel";
        public const string ErrorCodeProtoMetaFileName = "ErrorCodeExcel";
        public const string ErrorCodeFrameScriptFileName = "FrameErrorCode";
        public const string ErrorCodeBusinessScriptFileName = "BusinessErrorCode";
        public const string ErrorCodeFramePackageName = "DingFrame";
        public const string ErrorCodeBusinessPackageName = "Business.Data";
        public const string ErrorCodeScriptName = "ErrorCode";
        public const string ErrorCodeFrameSheetName = "Common";

        public readonly static Dictionary<string, string> EnumFixedField = new()
        {
            {"name", "string"}, 
            {"field", "string"},
            {"value", "string"},
            {"platform", "string"},
            {"comment", "string"}
        };
        public const char EnumFieldSplitSymbol = '|';

        public readonly static Dictionary<string, string> ErrorCodeFixedField = new()
        {
            {"code", "long"},
            {"codeStr", "string"},
            {"content", "%string"},
            {"comment", "string"}
        };
    }
}
