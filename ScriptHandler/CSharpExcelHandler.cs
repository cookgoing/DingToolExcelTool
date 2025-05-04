namespace DingToolExcelTool.ScriptHandler
{
    using Microsoft.CodeAnalysis;
    using Microsoft.CodeAnalysis.CSharp;
    using Microsoft.CodeAnalysis.Emit;
    using System.Data;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Collections.Concurrent;
    using Google.Protobuf;
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
        protected ConcurrentDictionary<string, Type> typeDic = new();
        protected ConcurrentDictionary<string, object> objDic = new();

        public void DynamicCompile(string[] csCodes)
        {
            typeDic.Clear();
            objDic.Clear();

            SyntaxTree[] syntaxTrees = new SyntaxTree[csCodes.Length];
            for (int i = 0; i < csCodes.Length; ++i) syntaxTrees[i] = CSharpSyntaxTree.ParseText(csCodes[i]);

            string pbDllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Google.Protobuf.dll");
            string runtimeDllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins", "System.Runtime.dll");
            string collectionDllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Plugins", "System.Collections.dll");
            MetadataReference[] systemReferences = new[]
            {
                MetadataReference.CreateFromFile(pbDllPath),
                MetadataReference.CreateFromFile(runtimeDllPath),
                MetadataReference.CreateFromFile(collectionDllPath),
            };

            CSharpCompilation compilation = CSharpCompilation.Create(
                "ProtoGeneratedAssembly",
                syntaxTrees: syntaxTrees,
                references: systemReferences,
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
            string csFieldName = FieldNameInProtoCS(fieldName);
            FieldInfo fieldInfo = type.GetField(csFieldName, BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new Exception($"[SetScriptProperty] C#脚本[{scriptName}]中，没有这个字段：{csFieldName}; excel field name: {fieldName}");
            object value = ExcelType2ScriptType(typeStr, valueStr);
            fieldInfo.SetValue(obj, value);
        }

        public void AddScriptList(string scriptName, string fieldName, string typeStr, string valueStr)
        {
            if (!ExcelUtil.IsArrType(typeStr)) throw new Exception($"[AddScriptList] 逻辑错误 不是数组类型，不能使用这个方法");

            string elementType = typeStr.Substring(0, typeStr.Length - 2);
            object value = ExcelType2ScriptType(elementType, valueStr);

            var (type, obj) = GetTypeObj(scriptName);
            fieldName = FieldNameInProtoCS(fieldName);
            FieldInfo fieldInfo = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new Exception($"[AddScriptList] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object listObj = fieldInfo.GetValue(obj) ?? throw new Exception($"[AddScriptList] C#对象[{scriptName}]中没有这个字段： {fieldName}");
            MethodInfo addMethod = listObj.GetType().GetMethod("Add", [value.GetType()]) ?? throw new Exception($"[AddScriptList] {scriptName}.{fieldName}; 这个字段不是列表; listType: {listObj.GetType()}; value: {value.GetType()}");

            addMethod.Invoke(listObj, [value]);
        }

        public void AddScriptMap(string scriptName, string fieldName, string typeStr, string keyData, string valueData)
        {
            if (!ExcelUtil.IsMapType(typeStr)) throw new Exception($"[AddScriptMap] 逻辑错误 不是字典类型，不能使用这个方法");

            string innerTypes = typeStr.Substring(4, typeStr.Length - 5);
            string[] keyValue = innerTypes.Split(',');
            string kType = keyValue[0], vType = keyValue[1];
            object kValue = ExcelType2ScriptType(kType, keyData);
            object vValue = ExcelType2ScriptType(vType, valueData);

            var (type, obj) = GetTypeObj(scriptName);
            fieldName = FieldNameInProtoCS(fieldName);
            FieldInfo fieldInfo = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new Exception($"[AddScriptMap] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object dicObj = fieldInfo.GetValue(obj) ?? throw new Exception($"[AddScriptMap] C#对象[{scriptName}]中没有这个字段： {fieldName}");
            MethodInfo addMethod = dicObj.GetType().GetMethod("Add", [kValue.GetType(), vValue.GetType()]) ?? throw new Exception($"[AddScriptList] {scriptName}.{fieldName}; 这个字段不是字典");

            addMethod.Invoke(dicObj, [kValue, vValue]);
        }

        public void AddListScriptObj(string scriptName, string itemScriptName)
        {
            var (type, obj) = GetTypeObj(scriptName);
            var (itemType, objType) = GetTypeObj(itemScriptName);
            string fieldName = FieldNameInProtoCS(CommonExcelCfg.ProtoMetaListFieldName);

            FieldInfo listField = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new Exception($"[AddListScriptValue] C#脚本[{scriptName}]中，没有这个字段：{fieldName}");
            object listObj = listField.GetValue(obj) ?? throw new Exception($"[AddListScriptValue] C#对象[{scriptName}]中没有这个字段： {fieldName}");
            MethodInfo addMethod = listObj.GetType().GetMethod("Add", [objType.GetType()]) ?? throw new Exception($"[AddScriptList] {scriptName}.{fieldName}; 这个字段不是列表");

            addMethod.Invoke(listObj, [objType]);

            RemoveObj(itemScriptName);
        }

        public void SerializeObjInProto(string scriptName, string outputFilePath)
        {
            if (string.IsNullOrEmpty(outputFilePath)) throw new Exception($"[SerializeObjInProto] 序列化路径是空的");
            string dirPath = Path.GetDirectoryName(outputFilePath);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            var (type, obj) = GetTypeObj(scriptName);

            MethodInfo writeMethod = type.GetMethod("WriteTo", [typeof(CodedOutputStream)]) ?? throw new Exception($"[SerializeObjInProto] 类：{scriptName}，没有 WriteTo 方法，难道不是通过 proto 生成的？");
            using var fileStream = new FileStream(outputFilePath, FileMode.OpenOrCreate);
            using CodedOutputStream outputSteam = new (fileStream);

            writeMethod.Invoke(obj, [outputSteam]);
        }


        public async Task GenerateExcelScript(ExcelHeadInfo headInfo, string messageName, string excelScriptOutputFile, bool isClient)
        {
            if (headInfo == null) throw new Exception($"[GenerateExcelScript] headInfo == null");
            if (string.IsNullOrEmpty(excelScriptOutputFile)) throw new Exception($"[GenerateExcelScript] 没有 Excel Script 的输出路径");

            string dirPath = Path.GetDirectoryName(excelScriptOutputFile);
            if (!Directory.Exists(dirPath)) Directory.CreateDirectory(dirPath);

            PlatformType platform = isClient ? PlatformType.Client : PlatformType.Server;
            string scriptName = Path.GetFileNameWithoutExtension(excelScriptOutputFile);
            string dataFileName = $"{messageName}{GeneralCfg.ProtoDataFileSuffix}";
            using StreamWriter sw = new(excelScriptOutputFile);
            StringBuilder scriptSB = new();
            StringBuilder dicFieldSB = new();
            StringBuilder classificationActionSB = new();
            StringBuilder dataLoadSB = new();

            List<ExcelFieldInfo> unionKey = new List<ExcelFieldInfo>(headInfo.UnionKey.Count);
            foreach (ExcelFieldInfo keyField in headInfo.UnionKey)
            {
                if ((platform & keyField.Platform) == 0) continue;

                unionKey.Add(keyField);
            }

            foreach (ExcelFieldInfo keyField in headInfo.IndependentKey)
            {
                if ((platform & keyField.Platform) == 0) continue;

                dicFieldSB.Append($"public Dictionary<{ExcelType2ScriptTypeStr(keyField.Type)}, {messageName}> {keyField.Name}Dic{{get; private set;}} = new();");
            }
            
            if (unionKey.Count > 0)
            {
                bool onlyOneKey = unionKey.Count == 1;
                dicFieldSB.Append("public Dictionary<");
                if (!onlyOneKey) dicFieldSB.Append('(');
                foreach (ExcelFieldInfo keyField in unionKey)
                {
                    dicFieldSB.Append($"{ExcelType2ScriptTypeStr(keyField.Type)},");
                }
                dicFieldSB.Remove(dicFieldSB.Length - 1, 1);
                if (!onlyOneKey) dicFieldSB.Append(')');
                dicFieldSB.Append($", {messageName}> Dic{{get; private set;}} = new();");
            }

            foreach (ExcelFieldInfo keyField in headInfo.IndependentKey)
            {
                if ((platform & keyField.Platform) == 0) continue;

                classificationActionSB.Append($"{keyField.Name}Dic.Add(item.{PropertyNameInProtoCS(keyField.Name)}, item);");
            }
            if (unionKey.Count > 0)
            {
                if (classificationActionSB.Length > 0) classificationActionSB.AppendLine();

                bool onlyOneKey = unionKey.Count == 1;
                classificationActionSB.Append("Dic.Add(");
                if (!onlyOneKey) classificationActionSB.Append('(');
                foreach (ExcelFieldInfo keyField in unionKey)
                {
                    classificationActionSB.Append($"item.{PropertyNameInProtoCS(keyField.Name)},");
                }
                classificationActionSB.Remove(classificationActionSB.Length - 1, 1);
                if (!onlyOneKey) classificationActionSB.Append(')');
                classificationActionSB.Append($", item);");
            }

            if (isClient) dataLoadSB.Append(@"AssetLoadModule assetModule = ModuleCollector.GetModule<AssetLoadModule>();
            byte[] serializedData = assetModule.Load<TextAsset>(protoDataPath)?.bytes;");
            else dataLoadSB.Append("byte[] serializedData = File.ReadAllBytes(protoDataPath);");

            scriptSB.AppendLine(@$"
using System.IO;
using System.Linq;
using System.Collections.Generic;
using UnityEngine;
using Google.Protobuf;
using DingFrame.Module;
using DingFrame.Module.AssetLoader;

namespace {GeneralCfg.ProtoMetaPackageName}
{{
    public class {scriptName}
    {{
        public static {scriptName} Ins{{get; private set;}}

        public {messageName}[] Datas{{get; private set;}}
        {dicFieldSB}

        public static {scriptName} CreateIns()
        {{
            if (Ins != null) return Ins;

            Ins = new {scriptName}();
            Ins.ParseProto();
            Ins.GenerateKV();
            return Ins;
        }}

        public static void ReleaseIns() => Ins = null;

        private void ParseProto()
        {{
            string protoDataPath = Path.Combine(GameConfigure.ExcelProtoDataPath, ""{dataFileName}"");
            {dataLoadSB}
            {messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix} msgList = {messageName}{CommonExcelCfg.ProtoMetaListMessageNameSuffix}.Parser.ParseFrom(serializedData);
        
            Datas = msgList.{CommonExcelCfg.ProtoMetaListFieldName}.ToArray();
        }}

        private void GenerateKV()
        {{
        
            foreach({messageName} item in Datas)
            {{
                {classificationActionSB}
            }}
        }}
    }}
}}
");

            await sw.WriteAsync(scriptSB.ToString());
            sw.Flush();
        }


        public (Type type, object obj) GetTypeObj(string scriptName)
        {
            string fullScriptName = $"{GeneralCfg.ProtoMetaPackageName}.{scriptName}";
            if (!typeDic.TryGetValue(scriptName, out Type type))
            {
                type = assembly.GetType(fullScriptName) ?? throw new Exception($"[GenerateTypeObj] proto生成的C#程序集不存在 这个类型：{fullScriptName}");
                typeDic.TryAdd(scriptName, type);
            }
            if (!objDic.TryGetValue(scriptName, out object obj))
            {
                obj = Activator.CreateInstance(type) ?? throw new Exception($"[GenerateTypeObj] 无法生成实例：type: {type}");
                objDic.TryAdd(scriptName, obj);
            }

            return (type, obj);
        }

        public bool RemoveObj(string scriptName) => objDic.TryRemove(scriptName, out _);

        public object ExcelType2ScriptType(string typeStr, string valueStr)
        {
            if (ExcelUtil.IsBaseType(typeStr))
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
                    case "string": return valueStr ?? string.Empty;
                    default: throw new Exception($"[CSharpHandler] 存在不合法的基础类型：{typeStr}");
                }
            }
            else if (ExcelUtil.IsEnumType(typeStr))
            {
                string fullTypetName = $"{GeneralCfg.ProtoMetaPackageName}.{typeStr}";
                Type enumType = assembly.GetType(fullTypetName) ?? throw new Exception($"[CSharpHandler] 这个类型：{fullTypetName} 通过程序集：{assembly.FullName} 不能生成 Type");

                if (!enumType.IsEnum) throw new Exception($"[CSharpHandler] 这个类型：{enumType} 不是枚举类型");
                if (!Enum.TryParse(enumType, valueStr, true, out var enumValue)) throw new Exception($"{valueStr} 不能转换成这个枚举类型：{enumType}");

                return enumValue;
            }
            else throw new Exception($"[CSharpHandler] 未知的类型：{typeStr}");
        }

        public string FieldNameInProtoCS(string fieldName) => $"{NameConverter.ConvertToCamelCase(fieldName)}_";

        public string PropertyNameInProtoCS(string fieldName) => NameConverter.ConvertToPascalCase(fieldName);

        
    }
}
