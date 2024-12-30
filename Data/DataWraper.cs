using DingToolExcelTool.Utils;

namespace DingToolExcelTool.Data
{
    internal class DataWraper(CustomDataInfo data)
    {
        public class PlatformOutputInfo
        {
            private CustomDataInfo.PlatformOutputInfo platformOutputInfo;

            public ScriptTypeEn ScriptType { get => platformOutputInfo.ScriptType; set => platformOutputInfo.ScriptType = value; }
            public string ProtoMetaOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoMetaOutputDir); set => platformOutputInfo.ProtoMetaOutputDir = value; }
            public string ProtoScriptOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoScriptOutputDir); set => platformOutputInfo.ProtoScriptOutputDir = value; }
            public string ProtoDataOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ProtoDataOutputDir); set => platformOutputInfo.ProtoDataOutputDir = value; }
            public string ExcelScriptOutputDir { get => ExcelUtil.ParsePath(platformOutputInfo.ExcelScriptOutputDir); set => platformOutputInfo.ExcelScriptOutputDir = value; }
            public string ErrorCodeFrameDir { get => ExcelUtil.ParsePath(platformOutputInfo.ErrorCodeFrameDir); set => platformOutputInfo.ErrorCodeFrameDir = value; }
            public string ErrorCodeBusinessDir { get => ExcelUtil.ParsePath(platformOutputInfo.ErrorCodeBusinessDir); set => platformOutputInfo.ErrorCodeBusinessDir = value; }

            public PlatformOutputInfo(CustomDataInfo.PlatformOutputInfo platformOutputInfo) =>  this.platformOutputInfo = platformOutputInfo;
        }

        public CustomDataInfo Data { get; private set; } = data;

        public string ExcelInputRootDir { get => ExcelUtil.ParsePath(Data.ExcelInputRootDir); set => Data.ExcelInputRootDir = value; }
        public bool OutputClient { get => Data.OutputClient; set => Data.OutputClient = value; }
        public bool OutputServer { get => Data.OutputServer; set => Data.OutputServer = value; }
        public PlatformOutputInfo ClientOutputInfo = new(data.ClientOutputInfo);
        public PlatformOutputInfo ServerOutputInfo = new(data.ServerOutputInfo);

        public string PreHanleProgramFile { get => ExcelUtil.ParsePath(Data.PreHanleProgramFile); set => Data.PreHanleProgramFile = value; }
        public string PreHanleProgramArgument {get => Data.PreHanleProgramArgument; set => Data.PreHanleProgramArgument = value; }
        public string AftHanleProgramFile { get => ExcelUtil.ParsePath(Data.AftHanleProgramFile); set => Data.AftHanleProgramFile = value; }
        public string AftHanleProgramArgument { get => Data.AftHanleProgramArgument; set => Data.AftHanleProgramArgument = value; }
    }
}
