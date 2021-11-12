using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Editor;

namespace excel2json
{
    /// <summary>
    /// 根据表头，生成C#类定义数据结构
    /// 表头使用三行定义：字段名称、字段类型、注释
    /// </summary>
    class CSDefineGenerator
    {
        struct FieldDef
        {
            public string name;
            public string type;
            public string comment;
        }

        string mCode;

        public string code {
            get {
                return this.mCode;
            }
        }

        public CSDefineGenerator(string excelName, ExcelLoader excel, string excludePrefix, string setNamespace, bool outputClient = true, bool fileLowercase = true)
        {
            //-- 创建代码字符串
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("//");
            sb.Append("/* This is an auto-generated meta script. */\n/* if you want to edit it, please dont use the ScriptToExcel feature, which might cause unhandled error.*/\n");
            sb.AppendLine("// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称");
            sb.AppendLine("// 2. 表格约定：第一行是变量名称，第二行是变量类型");
            sb.AppendLine();
            sb.AppendFormat("// Generate From {0}.xlsx", excelName);
            sb.AppendLine();
            sb.AppendLine();

            sb.AppendLine("using System.Collections.Generic;");
            sb.AppendLine();
            sb.AppendLine($"namespace {setNamespace}");
            sb.AppendLine("{");


            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];
                sb.Append(_exportSheet(sheet, excludePrefix, outputClient, fileLowercase));
            }
            sb.AppendLine("}");
            sb.AppendLine();
            sb.AppendLine("// End of Auto Generated Code");

            mCode = sb.ToString();
        }

        private string _exportSheet(DataTable sheet, string excludePrefix, bool outputClient = true, bool fileLowercase = true)
        {
            if (sheet.Columns.Count < 0 || sheet.Rows.Count < 2)
                return "";

            string sheetName = sheet.TableName;
            if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                return "";

            // get field list
            List<FieldDef> fieldList = new List<FieldDef>();
            DataRow typeRow = sheet.Rows[0];
            DataRow commentRow = sheet.Rows[1];

            foreach (DataColumn column in sheet.Columns)
            {
                // 过滤掉包含指定前缀的列
                string columnName = column.ToString();
                if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
                    continue;
                if (outputClient)
                {
                    if (columnName.StartsWith(Config.ServerSignal)) // 忽略服务端标志位相关的
                        continue;

                    if (columnName.StartsWith(Config.ClientSignal))
                    {
                        columnName = columnName.Substring(2);
                    }
                }

                if (fileLowercase)
                {
                    columnName = columnName.ToLower();
                }

                FieldDef field;
                field.name = columnName;
                var fieldType = GetProperty(typeRow[column].ToString());
                Console.WriteLine(fieldType);
                field.type = fieldType;
                field.comment = commentRow[column].ToString();

                fieldList.Add(field);
            }

            // export as string
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("\tpublic class {0}\r\n\t{{", sheet.TableName);
            sb.AppendLine();

            foreach (FieldDef field in fieldList)
            {
                sb.AppendFormat("\t\tpublic {0} {1} {{ get; set; }} // {2}", field.type, field.name, field.comment);
                sb.AppendLine();
            }

            sb.Append("\t}");
            sb.AppendLine();
            return sb.ToString();
        }

        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mCode);
            }
        }
        
        //判断声明类型
        public static string GetProperty(string type)
        {
            type = type.Trim();
            string generic = "";
            if (type == "int" || type == "Int" || type == "INT")
                type = "int";
            else if (type == "int32" || type == "Int32" || type == "INT32")
                type = "int";
            else if (type == "float" || type == "Float" || type == "FLOAT")
                type = "float";
            else if (type == "bool" || type == "Bool" || type == "BOOL")
                type = "bool";
            else if (type == "short" || type == "Short" || type == "SHORT")
                type = "short";
            else if (type == "int16" || type == "Int16" || type == "INT16")
                type = "short";
            else if (type == "long" || type == "Long" || type == "LONG")
                type = "long";
            else if (type == "int64" || type == "Int64" || type == "INT64")
                type = "long";
            
            else if (type.StartsWith("enum") || type.StartsWith("Enum") || type.StartsWith("ENUM"))
            {
                type = type.Split('|').LastOrDefault().Trim();
            }
            else if (type.StartsWith("list") || type.StartsWith("List") || type.StartsWith("LIST"))
            {
                var tmpType = CheckValueType(type.Split('|').LastOrDefault().Trim());
                generic = "<" + tmpType + ">";
                type = "List";
            }
            else if (type.EndsWith("[]")) // int切片 int[] 对excel 定义go类型进行支持
            {
                var tmpType = CheckValueType(type.Split('[').FirstOrDefault().Trim());
                generic = "<" + tmpType + ">";
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
            else if (type.StartsWith("dic<") || type.StartsWith("map<") || (type.StartsWith("<") && type.EndsWith(">"))) // dic<int32, int32>  map<int32, int32>  <int32, int32>
            {
                string pair = type.Split('<').LastOrDefault().Trim();
                pair = pair.Split('>').FirstOrDefault().Trim();
                
                string keyStr = "string", valueStr = "string";
                keyStr = CheckValueType(pair.Split(',')[0].Trim());
                valueStr = CheckValueType(pair.Split(',')[1].Trim());
                generic = "<" + keyStr + ", " + valueStr + ">";
                type = "Dictionary";
            }
            else
                type = "string";

            return type + generic;
        }
        
        // 用于泛型类型判断
        public static string CheckValueType(string typeName)
        {
            string type = "string";
            if (typeName == "int" || typeName == "Int" || typeName == "INT")
                type = "int";
            else if (type == "int32" || type == "Int32" || type == "INT32")
                type = "int";
            else if (typeName == "float" || typeName == "Float" || typeName == "FLOAT")
                type = "float";
            else if (typeName == "bool" || typeName == "Bool" || typeName == "BOOL")
                type = "bool";
            else if (typeName == "short" || typeName == "Short" || typeName == "SHORT")
                type = "short";
            else if (typeName == "int16" || typeName == "Int16" || typeName == "INT16")
                type = "short";
            else if (typeName == "long" || typeName == "Long" || typeName == "LONG")
                type = "long";
            else if (typeName == "Int64" || typeName == "Long" || typeName == "LONG")
                type = "long";
            
            return type;
        }
    }
}
