namespace DingToolExcelTool.Data
{
    internal class SingleExcelFieldInfo
    {
        public string Name;
        public string Type;
        public PlatformType Platform;
        public string Comment;
        public string Value;
        public bool LocalizationTxt;
        public bool LocalizationImg;
        public int StartRowIdx;
        public int EndRowIdx;
    }

    internal class SingleExcelHeadInfo
    {
        public string ScriptName;
        public List<SingleExcelFieldInfo> Fields;

        public void Trim()
        {
            for (int i = Fields.Count - 1; i >= 0; --i)
            {
                SingleExcelFieldInfo field = Fields[i];
                if (!string.IsNullOrEmpty(field.Name)
                    && !string.IsNullOrEmpty(field.Type)) continue;

                Fields.RemoveAt(i);
            }
        }

        public void Sort() => Fields.Sort((a, b) => a.StartRowIdx - b.StartRowIdx);

        public int GetFieldIdx(int columnIdx)
        {
            for (int i = 0; i < Fields.Count; ++i)
            {
                SingleExcelFieldInfo field = Fields[i];
                if (columnIdx < field.StartRowIdx) return -1;
                if (columnIdx <= field.EndRowIdx) return i;
            }

            return -1;
        }
    }
}
