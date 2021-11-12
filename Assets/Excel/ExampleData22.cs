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
		public string ID; // 编号
		public string Name; // 名称
		public string AssetName; // 资源编号
		public int HP; // 血
		public int Attack; // 攻击
		public int Defence; // 防御
		public int Enemy; // 敌人
		public List<string> EnemyList; // ["f1"]
		public List<int> EnemyIdList; // intlist
	}
}

// End of Auto Generated Code
