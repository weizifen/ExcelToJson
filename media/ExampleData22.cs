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
		public string ID { get; set; } // 编号
		public string AssetName { get; set; } // 资源编号
		public int HP { get; set; } // 血
		public int Attack { get; set; } // 攻击
		public int Defence { get; set; } // 防御
		public int Enemy { get; set; } // 敌人
		public List<string> EnemyList { get; set; } // f1
		public List<int> EnemyIdList { get; set; } // intlist
		public Dictionary<string, string> FafDict { get; set; } // sdasd
		public bool Bb { get; set; } // bd
	}
}

// End of Auto Generated Code
