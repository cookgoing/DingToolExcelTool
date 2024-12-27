﻿namespace DingToolExcelTool.ScriptHandler
{
    using Microsoft.CodeAnalysis;
    using Microsoft.CodeAnalysis.CSharp;
    using Microsoft.CodeAnalysis.Emit;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using DingToolExcelTool.Configure;
    using DingToolExcelTool.Utils;
    using DingToolExcelTool.Data;

    internal class CSharpExcelHandler : Singleton<CSharpExcelHandler>, IScriptExcelHandler
    {
        public Dictionary<string, string> BaseType2ScriptMap { get; private set; } = new ()
        { 
            {"int", "int"},
            {"long", "long"},
            {"double", "double"},
            {"bool", "bool"},
            {"string", "string"},
        };

        public string Suffix => ".cs";

        protected Assembly assembly;
        protected Dictionary<string, Type> typeDic = new();
        protected Dictionary<string, object> objDic = new();

        public void DynamicCompile(string[] csCodes)
        {
            SyntaxTree[] syntaxTrees = new SyntaxTree[csCodes.Length];
            for (int i = 0; i < csCodes.Length; ++i) syntaxTrees[i] = CSharpSyntaxTree.ParseText(csCodes[i]);

            string pbDllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Google.Protobuf.dll");
            CSharpCompilation compilation = CSharpCompilation.Create(
                "ProtoGeneratedAssembly",
                syntaxTrees: syntaxTrees,
                references: [AssemblyMetadata.CreateFromFile(pbDllPath).GetReference()],
                options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

            using var memoryStream = new MemoryStream();
            EmitResult result = compilation.Emit(memoryStream);

            if (!result.Success)
            {
                StringBuilder sb = new();
                foreach (var diagnostic in result.Diagnostics) sb.AppendLine(diagnostic.ToString());

                throw new Exception(sb.ToString());
            }

            memoryStream.Seek(0, SeekOrigin.Begin);
            assembly = Assembly.Load(memoryStream.ToArray());
        }


        public string ExcelType2ScriptTypeStr(string typeStr)
        {
            if (ExcelUtil.IsTypeLocalizationTxt(typeStr) || ExcelUtil.IsTypeLocalizationImg(typeStr)) return "string";
            if (ExcelUtil.IsBaseType(typeStr)) return BaseType2ScriptMap[typeStr];
            if (ExcelUtil.IsArrType(typeStr))
            {
                string elementType = typeStr.Substring(0, typeStr.Length - 2);
                if (BaseType2ScriptMap.TryGetValue(elementType, out string csType)) return $"{csType}[]";
                else return $"{elementType}[]";
            }
            if (ExcelUtil.IsMapType(typeStr))
            {
                string innerTypes = typeStr.Substring(4, typeStr.Length - 5);
                string[] keyValue = innerTypes.Split(',');
                string kType = keyValue[0], vType = keyValue[1];

                if (BaseType2ScriptMap.TryGetValue(kType, out string? value)) kType = value;
                if (BaseType2ScriptMap.TryGetValue(vType, out value)) vType = value;

                return $"Dictionary<{kType},{vType}>";
            }
            if (ExcelUtil.IsEnumType(typeStr))
            {
                return $"{typeStr}";
            }
            return string.Empty;
        }
        

        public void GenerateProtoScript(string metaInputFile, string protoScriptOutDir)
        {
            if (!File.Exists(metaInputFile)) throw new FileNotFoundException("[CSharpHandler] 目标 proto 文件未找到", metaInputFile);
            if (string.IsNullOrEmpty(protoScriptOutDir)) throw new Exception($"[SerializeObjInProto] proto script 的输出路径是空的");
            if (!Directory.Exists(protoScriptOutDir)) Directory.CreateDirectory(protoScriptOutDir);

            string arguments = $"--proto_path={Path.GetDirectoryName(metaInputFile)} " +
                               $"--csharp_out={protoScriptOutDir} " +
                               $"{Path.GetFileName(metaInputFile)}";

            ExcelUtil.GenerateProtoScript(arguments);
        }

        public void GenerateProtoScriptBatchly(string metaInputDir, string protoScriptOutDir)
        {
            if (!Directory.Exists(metaInputDir)) throw new FileNotFoundException("【GenerateCSCodeForProtoDir】目标 proto 文件夹未找到", metaInputDir);
            if (!Directory.Exists(protoScriptOutDir)) Directory.CreateDirectory(protoScriptOutDir);

            string[] protoFiles = Directory.GetFiles(metaInputDir, $"{GeneralCfg.ProtoMetaFileSuffix}");
            string arguments = $"--proto_path={metaInputDir} " +
                               $"--csharp_out={protoScriptOutDir} " +
                               string.Join(" ", protoFiles.Select(Path.GetFileName));

            ExcelUtil.GenerateProtoScript(arguments);
        }

        
        public void SetScriptValue(string scriptName, string fieldName, string typeStr, string valueStr)
        {
            if (ExcelUtil.IsArrType(typeStr) || ExcelUtil.IsMapType(typeStr)) throw new Exception($"[SetScriptProperty] 逻辑错误 要修改的数据是数组或者字典，不能使用这个方法");

            var (type, obj) = GetTypeObj(scriptName);

            PropertyInfo propertyInfo = type.GetProperty(fieldName) ?? throw new Exception($"[SetScriptProperty] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object value = ExcelType2ScriptType(typeStr, valueStr);
            propertyInfo.SetValue(obj, value);
        }

        public void AddScriptList(string scriptName, string fieldName, string typeStr, string valueStr)
        {
            if (!ExcelUtil.IsArrType(typeStr)) throw new Exception($"[AddScriptList] 逻辑错误 不是数组类型，不能使用这个方法");

            var (type, obj) = GetTypeObj(scriptName);
            PropertyInfo propertyInfo = type.GetProperty(fieldName) ?? throw new Exception($"[AddScriptList] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object listObj = propertyInfo.GetValue(obj) ?? throw new Exception($"[AddScriptList] C#对象[{scriptName}]中没有这个字段： {fieldName}");
            MethodInfo addMethod = listObj.GetType().GetMethod("Add") ?? throw new Exception($"[AddScriptList] {scriptName}.{fieldName}; 这个字段不是列表");

            string elementType = typeStr.Substring(0, typeStr.Length - 2);
            object value = ExcelType2ScriptType(elementType, valueStr);

            addMethod.Invoke(listObj, [value]);
        }

        public void AddScriptMap(string scriptName, string fieldName, string typeStr, string keyData, string valueData)
        {
            if (!ExcelUtil.IsMapType(typeStr)) throw new Exception($"[AddScriptMap] 逻辑错误 不是字典类型，不能使用这个方法");

            var (type, obj) = GetTypeObj(scriptName);
            PropertyInfo propertyInfo = type.GetProperty(fieldName) ?? throw new Exception($"[AddScriptMap] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object dicObj = propertyInfo.GetValue(obj) ?? throw new Exception($"[AddScriptMap] C#对象[{scriptName}]中没有这个字段： {fieldName}");
            MethodInfo addMethod = dicObj.GetType().GetMethod("Add") ?? throw new Exception($"[AddScriptList] {scriptName}.{fieldName}; 这个字段不是字典");

            string innerTypes = typeStr.Substring(4, typeStr.Length - 5);
            string[] keyValue = innerTypes.Split(',');
            string kType = keyValue[0], vType = keyValue[1];
            object kValue = ExcelType2ScriptType(kType, keyData);
            object vValue = ExcelType2ScriptType(vType, valueData);

            addMethod.Invoke(dicObj, [kValue, vValue]);
        }

        public void AddListScriptObj(string scriptName, string itemScriptName)
        {
            var (type, obj) = GetTypeObj(scriptName);
            var (itemType, objType) = GetTypeObj(itemScriptName);
            PropertyInfo listProperty = type.GetProperty(CommonExcelCfg.ProtoMetaListFieldName) ?? throw new Exception($"[AddListScriptValue] C#脚本[{scriptName}]中，没有这个字段：{CommonExcelCfg.ProtoMetaListFieldName}");
            object listObj = listProperty.GetValue(obj) ?? throw new Exception($"[AddListScriptValue] C#对象[{scriptName}]中没有这个字段： {CommonExcelCfg.ProtoMetaListFieldName}");
            MethodInfo addMethod = listObj.GetType().GetMethod("Add") ?? throw new Exception($"[AddScriptList] {scriptName}.{CommonExcelCfg.ProtoMetaListFieldName}; 这个字段不是列表");

            addMethod.Invoke(listObj, [objType]);
        }

        public void SerializeObjInProto(string scriptName, string outputFilePath)
        {
            if (string.IsNullOrEmpty(outputFilePath)) throw new Exception($"[SerializeObjInProto] 序列化路径是空的");
            string dirPath = Path.GetDirectoryName(outputFilePath);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            var (type, obj) = GetTypeObj(scriptName);

            MethodInfo writeMethod = type.GetMethod("WriteTo", [typeof(Stream)]) ?? throw new Exception($"[SerializeObjInProto] 类：{scriptName}，没有 WriteTo 方法，难道不是通过 proto 生成的？");
            using var fileStream = new FileStream(outputFilePath, FileMode.OpenOrCreate);

            writeMethod.Invoke(obj, [fileStream]);
        }


        public void GenerateExcelScript(ExcelHeadInfo headInfo, string excelScriptOutputFile, bool isClient)
        {
            if (headInfo == null) throw new Exception($"[GenerateExcelScript] headInfo == null");
            if (string.IsNullOrEmpty(excelScriptOutputFile)) throw new Exception($"[GenerateExcelScript] 没有 Excel Script 的输出路径");

            string dirPath = Path.GetDirectoryName(excelScriptOutputFile);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            string messageName = headInfo.MessageName;
            string scriptName = Path.GetFileNameWithoutExtension(excelScriptOutputFile);
            string dataFileName = $"{messageName}{GeneralCfg.ProtoDataFileSuffix}";
            using StreamWriter sw = new(excelScriptOutputFile);
            StringBuilder scriptSB = new();
            StringBuilder dicFieldSB = new();
            StringBuilder classificationActionSB = new();
            StringBuilder dataLoadSB = new();

            foreach (ExcelFieldInfo keyField in headInfo.IndependentKey)
            {
                dicFieldSB.Append($"public Dictionary<{ExcelType2ScriptTypeStr(keyField.Type)}, {messageName}> {keyField.Name}Dic{{get; private set;}}");
            }
            if (headInfo.UnionKey.Count > 0)
            {
                bool onlyOneKey = headInfo.UnionKey.Count == 1;
                dicFieldSB.Append("public Dictionary<");
                if (!onlyOneKey) dicFieldSB.Append('(');
                foreach (ExcelFieldInfo keyField in headInfo.UnionKey)
                {
                    dicFieldSB.Append($"{ExcelType2ScriptTypeStr(keyField.Type)},");
                }
                dicFieldSB.Remove(dicFieldSB.Length - 1, 1);
                if (!onlyOneKey) dicFieldSB.Append(')');
                dicFieldSB.Append($", {messageName}> Dic{{get; private set;}}");
            }

            foreach (ExcelFieldInfo keyField in headInfo.IndependentKey)
            {
                classificationActionSB.AppendLine($"{keyField.Name}Dic.Add(item.{keyField.Name}, item)");
            }
            if (headInfo.UnionKey.Count > 0)
            {
                bool onlyOneKey = headInfo.UnionKey.Count == 1;
                classificationActionSB.Append("Dic.Add(");
                if (!onlyOneKey) classificationActionSB.Append('(');
                foreach (ExcelFieldInfo keyField in headInfo.UnionKey)
                {
                    classificationActionSB.Append($"item.{keyField.Name},");
                }
                classificationActionSB.Remove(classificationActionSB.Length - 1, 1);
                if (!onlyOneKey) classificationActionSB.Append(')');
                classificationActionSB.Append($", item);");
            }

            if (isClient) dataLoadSB.Append(@"assetModule = ModuleCollector.GetModule<AssetLoadModule>();
byte[] serializedData = assetModule.Load<TextAsset>(protoDataPath)?.bytes;");
            else dataLoadSB.Append("byte[] serializedData = File.ReadAllBytes(protoDataPath);");

            scriptSB.AppendLine(@$"
using System.IO
using System.Collections.Generic;
using Google.Protobuf;
using DingFrame.Module.AssetLoader;

namespace {GeneralCfg.ProtoMetaPackageName};
public class {scriptName}
{{
    public static {scriptName} Ins{{get; private set;}}

    public {messageName}[] Datas{{get; private set;}}
    {dicFieldSB}

    public static {scriptName} CreateIns()
    {{
        Ins = new {scriptName}();
        Ins.ParseProto();
        Ins.GenerateKV();
        return Ins;
    }}

    public static void ReleaseIns() => Ins = null;

    private void ParseProto()
    {{
        string protoDataPath = Path.Combine(GameConfigure.ExcelDataPath, {dataFileName});
        {dataLoadSB}
        {messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix} msgList = {messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}.Parser.ParseFrom(serializedData);
        
        this.Datas = msgList.{CommonExcelCfg.ProtoMetaListFieldName}.ToArray();
    }}

    private void GenerateKV()
    {{
        
        foreach({messageName} item in Datas)
        {{
            {classificationActionSB}
        }}
    }}
}}
");

            sw.Write(scriptSB.ToString());
            sw.Flush();
        }


        private (Type type, object obj) GetTypeObj(string scriptName)
        {
            if (!typeDic.TryGetValue(scriptName, out Type type))
            {
                type = assembly.GetType(scriptName) ?? throw new Exception($"[GenerateTypeObj] proto生成的C#程序集不存在 这个类型：{scriptName}");
                typeDic.Add(scriptName, type);
            }
            if (!objDic.TryGetValue(scriptName, out object obj))
            {
                obj = Activator.CreateInstance(type) ?? throw new Exception($"[GenerateTypeObj] 无法生成实例：type: {type}");
                objDic.Add(scriptName, obj);
            } 

            return (type, obj);
        }

        private object ExcelType2ScriptType(string typeStr, string valueStr)
        {
            if (ExcelUtil.IsTypeLocalizationTxt(typeStr) || ExcelUtil.IsTypeLocalizationImg(typeStr)) return valueStr;
            else if (ExcelUtil.IsBaseType(typeStr))
            {
                switch (typeStr)
                {
                    case "int":
                        if (!int.TryParse(valueStr, out int intValue)) throw new Exception($"[CSharpHandler] 内容不能转换。 type: {typeStr}; value: {valueStr}");

                        return intValue;
                    case "long":
                        if (!long.TryParse(valueStr, out long longValue)) throw new Exception($"[CSharpHandler] 内容不能转换。 type: {typeStr}; value: {valueStr}");

                        return longValue;
                    case "double":
                        if (!double.TryParse(valueStr, out double doubleValue)) throw new Exception($"[CSharpHandler] 内容不能转换。 type: {typeStr}; value: {valueStr}");

                        return doubleValue;
                    case "bool":
                        if (!bool.TryParse(valueStr, out bool boolValue)) throw new Exception($"[CSharpHandler] 内容不能转换。 type: {typeStr}; value: {valueStr}");

                        return boolValue;
                    case "string": return valueStr;
                    default: throw new Exception($"[CSharpHandler] 存在不合法的基础类型：{typeStr}");
                }
            }
            else if (ExcelUtil.IsEnumType(typeStr))
            {
                Type enumType = assembly.GetType(typeStr) ?? throw new Exception($"[CSharpHandler] 这个类型：{typeStr} 通过程序集：{assembly.FullName} 不能生成 Type");

                if (!enumType.IsEnum) throw new Exception($"[CSharpHandler] 这个类型：{enumType} 不是枚举类型");
                if (!Enum.TryParse(enumType, valueStr, true, out var enumValue)) throw new Exception($"{valueStr} 不能转换成这个枚举类型：{enumType}");

                return enumValue;
            }
            else throw new Exception($"[CSharpHandler] 未知的类型：{typeStr}");
        }
    }
}