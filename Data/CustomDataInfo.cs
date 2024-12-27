namespace DingToolExcelTool.Data
{
    internal class CustomDataInfo
    {
        public class PlatformOutputInfo
        {
            public ScriptTypeEn ScriptType;
            public string ProtoMetaOutputDir;
            public string ProtoScriptOutputDir;
            public string ProtoDataOutputDir;
            public string ExcelScriptOutputDir;
            public string ErrorCodeFrameDir;
            public string ErrorCodeBusinessDir;
        }

        public string ExcelInputRootDir;
        public bool OutputClient;
        public bool OutputServer;
        public PlatformOutputInfo ClientOutputInfo;
        public PlatformOutputInfo ServerOutputInfo;

        public string PreHanleProgramFile;
        public string PreHanleProgramArgument;

        public string AftHanleProgramFile;
        public string AftHanleProgramArgument;
    }
}
