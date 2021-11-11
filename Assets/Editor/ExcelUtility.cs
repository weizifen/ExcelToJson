using UnityEngine;
using System.Collections;
using System.Collections.Generic;
using Excel;
using System.Data;
using System.IO;
using System.Text;
using System.Reflection;
using System.Reflection.Emit;
using System;
using Newtonsoft.Json;

public partial class ExcelUtility
{

	/// <summary>
	/// 表格数据集合
	/// </summary>
	private DataSet mResultSet;

	/// <summary>
	/// 构造函数
	/// </summary>
	/// <param name="excelFile">Excel file.</param>
	public ExcelUtility (string excelFile)
	{
		FileStream mStream = File.Open (excelFile, FileMode.Open, FileAccess.Read);
		IExcelDataReader mExcelReader = ExcelReaderFactory.CreateOpenXmlReader (mStream);
		mResultSet = mExcelReader.AsDataSet ();
	}
	

	/// <summary>
	/// 转换为Json
	/// </summary>
	/// <param name="JsonPath">Json文件路径</param>
	/// <param name="Header">表头行数</param>
	public void ConvertToJson (string JsonPath, Encoding encoding, bool isDesc, string outputCs = "")
	{
		//判断Excel文件中是否存在数据表
		if (mResultSet.Tables.Count < 1)
			return;

		//默认读取第一个数据表
		DataTable mSheet = mResultSet.Tables [0];

		//判断数据表内是否存在数据
		if (mSheet.Rows.Count < 1)
			return;

		//读取数据表行数和列数
		int rowCount = mSheet.Rows.Count;
		int colCount = mSheet.Columns.Count;

		//准备一个列表存储整个表的数据
		List<Dictionary<string, object>> table = new List<Dictionary<string, object>> ();

		
		var readIndex = 4;
		if (!isDesc)
		{
			readIndex = 3;
		}

		//读取数据
		for (int i = readIndex; i < rowCount; i++) {
			// 解决空白行
			if (mSheet.Rows[i][0].ToString().Length == 0)
			{
				continue;
			}
			//准备一个字典存储每一行的数据
			Dictionary<string, object> row = new Dictionary<string, object> ();
			for (int j = 0; j < colCount; j++) {
				//读取第1行数据作为表头字段
				string field = mSheet.Rows [0] [j].ToString ();
				
				var isEmpty = field.Contains("!");
				var isEmpty2 = field.Contains("#");

				if (isEmpty || isEmpty2 || field.Length == 0)
				{
					continue;
				}
				
				//Key-Value对应
				row [field] = mSheet.Rows [i] [j].ToString();
			}

			//添加到表数据中
			table.Add (row);
		}

		//生成Json字符串
		string json = JsonConvert.SerializeObject (table, Newtonsoft.Json.Formatting.Indented);
		
		// litjson 使用的是unicode  将excel 转成json会有问题 所以先不采用
		// JsonWriter jw = new JsonWriter();
		// jw.PrettyPrint = true;
		// JsonMapper.ToJson(table, jw);
		//写入文件
		using (FileStream fileStream=new FileStream(JsonPath,FileMode.Create,FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream, encoding)) {
				textWriter.Write (json);
			}
		}

		var output = "";
		if (outputCs == "")
		{
			output = JsonPath.Replace(".json",".cs");
		}
		else
		{
			output = outputCs;
		}
		
		var arr = output.Split('/');
		var csName = arr[arr.Length - 1].Replace(".cs", "");

		var ext = new ExcelMediumData();
		ext.excelName = csName;
		ext.propertyType = new Dictionary<string, string>();
		// ext.dataEachLine = new List<Dictionary<string, string>>();
		
		//读取数据
		for (int k = 0; k < colCount; k++) {
			//读取第2行数据作为表头字段
			string field = mSheet.Rows [0] [k].ToString ();
			//读取第3行数据作为类型
			string typeField = mSheet.Rows [2] [k].ToString ();
			ext.propertyType.Add(field, typeField);
		}

		
		var codeStr = CreateAssetCode(ext);
		//写入文件
		using (FileStream fileStream=new FileStream(output,FileMode.Create,FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream, encoding)) {
				textWriter.Write (codeStr);
			}
		}
	}
	
	
	

    /// <summary>
    /// 转换为CSV
    /// </summary>
    public void ConvertToCSV (string CSVPath, Encoding encoding, bool isDesc)
	{
		//判断Excel文件中是否存在数据表
		if (mResultSet.Tables.Count < 1)
			return;

		//默认读取第一个数据表
		DataTable mSheet = mResultSet.Tables [0];

		//判断数据表内是否存在数据
		if (mSheet.Rows.Count < 1)
			return;

		//读取数据表行数和列数
		int rowCount = mSheet.Rows.Count;
		int colCount = mSheet.Columns.Count;

		//创建一个StringBuilder存储数据
		StringBuilder stringBuilder = new StringBuilder ();

		List<int> tmpList = new List<int>();

		var readIndex = 1;
		if (!isDesc)
		{
			readIndex = 0;
		}

		//读取数据
		for (int i = readIndex; i < rowCount; i++) {
			// 解决空白行
			if (mSheet.Rows[i][0].ToString().Length == 0)
			{
				continue;
			}

			for (int j = 0; j < colCount; j++) {
				//使用","分割每一个数值

				if (i == readIndex)
				{
					var tKey = mSheet.Rows[i][j].ToString();
					var isEmpty = tKey.Contains("!");
					var isEmpty2 = tKey.Contains("#");

					if (isEmpty || isEmpty2 || tKey.Length == 0)
					{
						tmpList.Add(j);
						continue;
					}
				}

				if (tmpList.Contains(j))
				{
					continue;
				}
				var value = mSheet.Rows[i][j].ToString();
				if (value.Contains(","))
				{
					// 需转义 不然中间有逗号的字符串读起来会有问题
					value = $"\"{value}\"";
					stringBuilder.Append (value + ",");
				}
				else
				{
					stringBuilder.Append (mSheet.Rows [i] [j] + ",");
				}
			}

			stringBuilder.Remove(stringBuilder.Length - 1, 1);
			//使用换行符分割每一行
			stringBuilder.Append ("\r\n");	
		}

		//写入文件
		using (FileStream fileStream = new FileStream(CSVPath, FileMode.Create, FileAccess.Write)) {
			using (TextWriter textWriter = new StreamWriter(fileStream, encoding)) {
				textWriter.Write (stringBuilder.ToString ());
			}
		}
	}
    
}

