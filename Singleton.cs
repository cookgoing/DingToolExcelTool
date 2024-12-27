namespace DingToolExcelTool
{
    internal class Singleton<T> where T : Singleton<T>, new()
    {
        protected static T _instance;
        public static T Instance { get => _instance ??= CreateInstance(); }

        public static T CreateInstance()
        {
            if (_instance != null) return _instance;

            _instance = new T();
            return _instance;
        }
    }
}
