//
/* This is an auto-generated meta script. */
/* if you want to edit it, please dont use the ScriptToExcel feature, which might cause unhandled error.*/
// 1. 每个 Sheet 形成一个 Struct 定义, Sheet 的名称作为 Struct 的名称
// 2. 表格约定：第一行是变量名称，第二行是变量类型

// Generate From ExampleData22.xlsx

using System.Collections.Generic;

namespace HotUpdateScripts.Xiuxian
{
	public class NPC
	{
		public string id { get; set; } // 编号
		public int hp { get; set; } // 血
		public int attack { get; set; } // 攻击
		public int defence { get; set; } // 防御
		public int enemy { get; set; } // 敌人
		public List<string> enemylist { get; set; } // f1
		public List<int> enemyidlist { get; set; } // intlist
		public Dictionary<string, string> fafdict { get; set; } // sdasd
		public bool bb { get; set; } // bd
		public short enemy2 { get; set; } // dsad
		public ItemTypeEnum testenum { get; set; } // 枚举测试
		public List<string> finaltest1 { get; set; } // int32[]类型测试
		public Dictionary<string, string> finaltest2 { get; set; } // <string, string>类型测试
		public Dictionary<string, string> finaltest3 { get; set; } // dic<string, string>类型测试
		public int finaltest4 { get; set; } // int32类型测试
	}
}

// End of Auto Generated Code
