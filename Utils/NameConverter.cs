namespace DingToolExcelTool.Utils
{
    using System.IO;
    using System.Linq;
    using System.Diagnostics;
    using ClosedXML.Excel;
    using DingToolExcelTool.Data;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.ExcelHandler;
    using DingToolExcelTool.ScriptHandler;
    using System.Globalization;
    using System.Text;

    internal static class NameConverter
    {
        /// <summary>
        /// 将 snake_case 字符串转换为 PascalCase。
        /// </summary>
        public static string ConvertToPascalCase(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var words = input.Split('_');
            var result = new StringBuilder();

            foreach (var word in words)
            {
                if (string.IsNullOrEmpty(word))
                    continue;

                result.Append(char.ToUpperInvariant(word[0]) + word.Substring(1));
            }

            return result.ToString();
        }

        /// <summary>
        /// 将 snake_case 字符串转换为 camelCase。
        /// </summary>
        public static string ConvertToCamelCase(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            var pascalCase = ConvertToPascalCase(input);
            if (pascalCase.Length == 0)
                return pascalCase;

            return char.ToLowerInvariant(pascalCase[0]) + pascalCase.Substring(1);
        }
    }
}
