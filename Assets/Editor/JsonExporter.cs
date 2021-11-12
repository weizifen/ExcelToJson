using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Editor;
using Newtonsoft.Json;

namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter
    {
        string mContext = "";
        int mHeaderRows = 0;

        public string context {
            get {
                return mContext;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>
        public JsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, bool forceSheetName, int headerRows, string excludePrefix, bool cellJson, bool allString, bool outputClient)
        {
            mHeaderRows = headerRows - 1;
            List<DataTable> validSheets = new List<DataTable>();
            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];

                // 过滤掉包含特定前缀的表单
                string sheetName = sheet.TableName;
                if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                    continue;

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheet);
            }

            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            if (!forceSheetName && validSheets.Count == 1)
            {   // single sheet

                //-- convert to object
                object sheetValue = convertSheet(validSheets[0], exportArray, lowcase, excludePrefix, cellJson, allString, outputClient);
                //-- convert to json string
                mContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
            }
            else
            { // mutiple sheet

                Dictionary<string, object> data = new Dictionary<string, object>();
                foreach (var sheet in validSheets)
                {
                    object sheetValue = convertSheet(sheet, exportArray, lowcase, excludePrefix, cellJson, allString, outputClient);
                    data.Add(sheet.TableName, sheetValue);
                }

                //-- convert to json string
                mContext = JsonConvert.SerializeObject(data, jsonSettings);
            }
        }

        private Dictionary<int, PropertyInfo> column2Property;
        private Type classType;
        
        private object convertSheet(DataTable sheet, bool exportArray, bool lowcase, string excludePrefix, bool cellJson, bool allString, bool outputClient)
        {
            
            #region 第一行 记录一下字段
            DataRow oneRow = sheet.Rows[0];
            // get field list
            var RuntimeAssembly = AppDomain.CurrentDomain.GetAssemblies().First(Assembly => Assembly.FullName.StartsWith("Assembly-CSharp"));
            classType = RuntimeAssembly.GetTypes().First(type =>
            {
                var tmpArr = type.Name.Split('.');
                var str = tmpArr[tmpArr.Length - 1];
                return str == sheet.TableName;
            });
            column2Property = new Dictionary<int, PropertyInfo>();
            DataRow typeRow = sheet.Rows[0];
            DataRow commentRow = sheet.Rows[1];

            var index = 0;
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
            
                var property = classType.GetProperty(columnName, BindingFlags.Instance | BindingFlags.Public);
                if (property != null)
                    column2Property[index] = property;
                index += 1;
            }
            #endregion
            if (exportArray)
                return convertSheetToArray(sheet, lowcase, excludePrefix, cellJson, allString, outputClient);
            else
                return convertSheetToDict(sheet, lowcase, excludePrefix, cellJson, allString, outputClient);
        }
        private object convertSheetToArray(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString, bool outputClient)
        {
            List<object> values = new List<object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson, allString, outputClient)
                    );
            }

            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson, bool allString, bool outputClient)
        {
            Dictionary<string, object> importData =
                new Dictionary<string, object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson, allString, outputClient);
                // 多余的字段
                // rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row, bool lowcase, int firstDataRow, string excludePrefix, bool cellJson, bool allString, bool outputClient)
        {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            object targetClass = Activator.CreateInstance(classType);
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

                object value = row[column];

                // 尝试将单元格字符串转换成 Json Array 或者 Json Object
                if (cellJson)
                {
                    string cellText = value.ToString().Trim();
                    if (cellText.StartsWith("[") || cellText.StartsWith("{"))
                    {
                        try
                        {
                            object cellJsonObj = JsonConvert.DeserializeObject(cellText);
                            if (cellJsonObj != null)
                                value = cellJsonObj;
                        }
                        catch (Exception exp)
                        {
                        }
                    }
                }
                
                // Console.WriteLine(value.GetType().Name);
                // Console.WriteLine(row[column].ToString());
                if (value.GetType() == typeof(System.DBNull))
                {
                    value = getColumnDefault(sheet, column, firstDataRow);
                }
                else if (value.GetType() == typeof(double))
                { // 去掉数值字段的“.0”
                    double num = (double)value;
                    if ((int)num == num)
                        value = (int)num;
                }

                //全部转换为string
                //方便LitJson.JsonMapper.ToObject<List<Dictionary<string, string>>>(textAsset.text)等使用方式 之后根据自己的需求进行解析
                if (allString && !(value is string))
                {
                    value = value.ToString();
                }

                string fieldName = columnName;
                // 表头自动转换成小写
                if (lowcase)
                    fieldName = fieldName.ToLower();

                if (string.IsNullOrEmpty(fieldName))
                    fieldName = string.Format("col_{0}", col);
                
                var propertyInfo = column2Property[col];
                object realValue = ParseStr(value.ToString(), propertyInfo.PropertyType);
                
                column2Property[col].SetValue(targetClass, realValue);
                
                rowData[fieldName] = realValue;
                col++;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull))
                {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mContext);
            }
        }

        #region 解析获得真实值

        private static object ParseStr(string strValue, Type propertyType)
        {
            if (string.IsNullOrEmpty(strValue)) strValue = String.Empty;
            strValue = strValue.TrimEnd();
            strValue = strValue.TrimStart();
            //如果有一行的第一列是空，则直接跳出
            object realValue = strValue;
            //根据类型不同填充属性
            //值为空
            if (string.IsNullOrEmpty(strValue) && (propertyType.IsValueType || propertyType == typeof(string)))
            {
                //如果是指类型则设置默认值
                if (propertyType.IsValueType)
                {
                    realValue = Activator.CreateInstance(propertyType);
                }

                if (propertyType == typeof(string))
                {
                    realValue = string.Empty;
                }
            }
            //如果是list
            else if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(List<>))
            {
                IList list = (IList)Activator.CreateInstance(propertyType);
                if (!string.IsNullOrEmpty(strValue))
                {
                    var values = strValue.Split('|');

                    var genericType = propertyType.GenericTypeArguments[0];
                    foreach (string s in values)
                    {
                        list.Add(ParseStr(s, genericType));
                    }
                }

                realValue = list;
            }
            //如果是字典
            else if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Dictionary<,>))
            {
                IDictionary dic = (IDictionary)Activator.CreateInstance(propertyType);
                if (!string.IsNullOrEmpty(strValue))
                {
                    var values = strValue.Split('|');
                    var keyType = propertyType.GenericTypeArguments[0];
                    var valueType = propertyType.GenericTypeArguments[1];
                    foreach (var s in values)
                    {
                        var keyValues = s.Split(':');
                        var key = ParseStr(keyValues[0], keyType);
                        var val = ParseStr(keyValues[1], valueType);
                        dic.Add(key, val);
                    }
                }

                realValue = dic;
            }
            //如果是枚举
            else if (propertyType.IsEnum)
            {
                realValue = Enum.Parse(propertyType, strValue);
            }
            //基础类型
            else
            {
                realValue = Convert.ChangeType(strValue, propertyType);
            }

            return realValue;
        }

        #endregion
    }
}
