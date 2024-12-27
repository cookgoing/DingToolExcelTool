namespace DingToolExcelTool.Data
{
    internal class EnumFieldInfo
    {
        public string Name;
        public int Value;
    }

    internal class EnumInfo
    {
        public string Name;
        public EnumFieldInfo[] Fields;
        public PlatformType Platform;
        public string Comment;
    }
}
