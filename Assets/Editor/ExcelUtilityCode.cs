using System.Collections.Generic;
using System.Linq;
using System.Text;

public class ExcelMediumData
{
    public string excelName;

    //Dictionary<字段名称, 字段类型>
    public Dictionary<string, string> propertyType;

    //Dictionary<字段名称, 字段中文提示>
    // public Dictionary<string, string> propertyChinese;

    //List<行数据>，List<Dictionary<字段名称, 一行的每个单元格字段值>>
    // public List<Dictionary<string, string>> dataEachLine;
}

public partial class ExcelUtility
{
    // 生成数据代码
    public static string CreateAssetCode(ExcelMediumData excelData)
    {
        if (excelData == null)
            return null;
        var dataName = excelData.excelName + "Data";
        if (string.IsNullOrEmpty(dataName))
            return null;
        Dictionary<string, string> propertyType = excelData.propertyType;
        if (propertyType == null || propertyType.Count == 0)
            return null;
        // if (excelData.dataEachLine == null || excelData.dataEachLine.Count == 0)
        //     return null;
        
        
        StringBuilder classSource = new StringBuilder();
        classSource.Append("/* This is an auto-generated meta script. */\n/* if you want to edit it, please dont use the ScriptToExcel feature, which might cause unhandled error.*/\n");
        classSource.Append("\n");
        classSource.Append("using UnityEngine;\n");
        classSource.Append("using System.Collections.Generic;\n");
        classSource.Append("using System.Linq;\n");
        classSource.Append("using System.IO;\n");
        classSource.Append("using UnityEditor;\n");
        // classSource.Append("using Sirenix.OdinInspector;\n");
        classSource.Append("\n");
        classSource.Append(CreateDataClass(dataName, propertyType));
        classSource.Append("\n");
        return classSource.ToString();
    }

    //数据赋值代码
    private static StringBuilder CreateDataClass(string dataName, Dictionary<string, string> propertyType)
    {
        var nameSpace = "namespace HotUpdateScripts.Xiuxian";
        StringBuilder classSource = new StringBuilder();
        classSource.Append(nameSpace);
        classSource.Append("{\n");
        classSource.Append("\t[System.Serializable]\n");
        classSource.Append("\tpublic class " + dataName + "\n");
        classSource.Append("\t{\n");
        classSource.Append("\t\t#region --- Auto Config --- \n");
        foreach (var item in propertyType)
        {
            classSource.Append(CreateCodeProperty(item.Key, item.Value));
        }

        classSource.Append("\t\t#endregion \n");
        classSource.Append("\t}\n");
        classSource.Append("}\n");
        return classSource;
    }


    //判断声明类型
    private static string CreateCodeProperty(string name, string type)
    {
        if (string.IsNullOrEmpty(name))
            return null;
        if (name == "idName")
            return null;
        type = type.Trim();
        string generic = "";
        if (type == "int" || type == "Int" || type == "INT")
            type = "int";
        else if (type == "float" || type == "Float" || type == "FLOAT")
            type = "float";
        else if (type == "bool" || type == "Bool" || type == "BOOL")
            type = "bool";
        else if (type.StartsWith("enum") || type.StartsWith("Enum") || type.StartsWith("ENUM"))
        {
            type = type.Split('|').LastOrDefault().Trim();
        }
        else if (type.StartsWith("list") || type.StartsWith("List") || type.StartsWith("LIST"))
        {
            generic = "<" + type.Split('|').LastOrDefault().Trim() + ">";
            type = "List";
        }
        else if (type.StartsWith("dictionary") || type.StartsWith("Dictionary") || type.StartsWith("DICTIONARY"))
        {
            string pair = type.Split('|').LastOrDefault().Trim();
            string keyStr = "string", valueStr = "string";
            keyStr = CheckValueType(pair.Split(',')[0].Trim());
            valueStr = CheckValueType(pair.Split(',')[1].Trim());
            generic = "<" + keyStr + ", " + valueStr + ">";
            type = "Dictionary";
        }
        else
            type = "string";

        string propertyStr = "\t\tpublic " + type + generic + " " + name + ";\n";
        return propertyStr;
    }

    // 用于泛型类型判断
    public static string CheckValueType(string typeName)
    {
        string type = "string";
        if (typeName == "int" || typeName == "Int" || typeName == "INT")
            type = "int";
        else if (typeName == "float" || typeName == "Float" || typeName == "FLOAT")
            type = "float";
        else if (typeName == "bool" || typeName == "Bool" || typeName == "BOOL")
            type = "bool";
        return type;
    }
}