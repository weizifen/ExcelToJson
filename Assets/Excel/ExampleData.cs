//
/* This is an auto-generated meta script. */
/* if you want to edit it, please dont use the ScriptToExcel feature, which might cause unhandled error.*/
// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称
// 2. 表格约定：第一行是变量名称，第二行是变量类型

// Generate From ExampleData.xlsx

using System.Collections.Generic;

namespace HotUpdateScripts.Xiuxian
{
	public class NPCFF
	{
		public string ID { get; set; } // 编号
		public string Name { get; set; } // 名称
		public string AssetName { get; set; } // 资源编号
		public int HP { get; set; } // 血
		public int Attack { get; set; } // 攻击
		public int Defence { get; set; } // 防御
		public List<string> DateTest { get; set; } // 测试日期
		public Dictionary<string, string> DateDict { get; set; } // 测试日期
	}
}

// End of Auto Generated Code
