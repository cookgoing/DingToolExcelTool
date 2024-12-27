namespace DingToolExcelTool.Data
{
    public enum LogType
    {
        Info,
        Warn,
        Error,
    }

    public enum HeadType
    {
        name,
        type,
        platform,
        comment
    }

    public enum KeyType
    { 
        No,
        Union,
        Independent,
    }

    internal enum ScriptTypeEn 
    { 
        CSharp,
    }

    internal enum PlatformType
    {
        Empty = 0,
        Client = 1,
        Server = 2,
        All = 3,
    }
}
