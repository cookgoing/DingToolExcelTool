namespace DingToolExcelTool.Data
{
    internal class ExcelFieldInfo
    {
        public string Name;
        public string Type;
        public PlatformType Platform;
        public string Comment;
        public bool LocalizationTxt;
        public bool LocalizationImg;
        public int StartColumnIdx;
        public int EndColumnIdx;
    }

    internal class ExcelHeadInfo
    {
        public string MessageName;
        public List<ExcelFieldInfo> Fields;
        public List<ExcelFieldInfo> UnionKey;
        public List<ExcelFieldInfo> IndependentKey;

        public void Sort() => Fields.Sort((a, b) => a.StartColumnIdx - b.StartColumnIdx);

        public int GetFieldIdx(int columnIdx)
        {
            for (int i = 0; i < Fields.Count; ++i)
            {
                ExcelFieldInfo field = Fields[i];
                if (columnIdx < field.StartColumnIdx) return -1;
                if (columnIdx < field.EndColumnIdx) return i;
            }

            return -1;
        }
    }
}
