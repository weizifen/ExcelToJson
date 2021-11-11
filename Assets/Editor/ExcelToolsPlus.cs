using System.Collections;
using System.Collections.Generic;
using System.Text;
using Sirenix.OdinInspector;
using Sirenix.OdinInspector.Editor;
using Sirenix.Utilities;
using Sirenix.Utilities.Editor;
using UnityEditor;
using UnityEngine;

namespace Editor
{
    public class ExcelToolsPlus : OdinEditorWindow
    {
        [Title("资源收集")] [AssetList(Path = "Excel", CustomFilterMethod = "IsExcel")]
        public List<Object> ExcelList;

        [Title("选择格式类型")] [HideLabel] [EnumPaging]
        public FormatType formatType = FormatType.JSON;

        [Title("选择编码类型")] [HideLabel] [ValueDropdown("codeTypeList")]
        public int codeType = 0;

        [Title("输出文件夹(json,csv), 默认(Assets/Excel)")] [HideLabel] [FolderPath]
        public string outPath = "Assets/Excel";

        [Title("输出文件夹(cs), 默认(Assets/Excel)")] [HideLabel] [FolderPath]
        public string outCsPath = "Assets/Excel";
        
        [Title("保留Excel源文件")] [ShowInInspector] [LabelText("保留源文件")]
        public bool keepSource = true;

        [Button(ButtonSizes.Gigantic)]
        [LabelText("执行")]
        public void Exec()
        {
            foreach (Object obj in ExcelList)
            {
                var assetsPath = AssetDatabase.GetAssetPath(obj);
                //获取Excel文件的绝对路径
                string excelPath = pathRoot + "/" + assetsPath;
                //构造Excel工具类
                ExcelUtility excel = new ExcelUtility(excelPath);

                //判断编码类型
                Encoding encoding = null;
                if (codeType == 0)
                {
                    encoding = Encoding.GetEncoding("utf-8");
                }
                else if (codeType == 1)
                {
                    encoding = Encoding.GetEncoding("gb2312");
                }

                //判断输出类型
                string output = "";
                string outputCs = "";
                if (formatType == FormatType.JSON)
                {
                    output = pathRoot + "/" + outPath + "/" + obj.name + ".json";
                    outputCs = pathRoot + "/" + outCsPath + "/" + obj.name + ".cs";
                    
                    excel.ConvertToJson(output, encoding, true, outputCs);
                }
                else if (formatType == FormatType.CSV)
                {
                    output = excelPath.Replace(".xlsx", ".csv");
                    excel.ConvertToCSV(output, encoding, true);
                }

                //判断是否保留源文件
                if (!keepSource)
                {
                    FileUtil.DeleteFileOrDirectory(excelPath);
                }

                //刷新本地资源
                AssetDatabase.Refresh();
                
                Close();
            }

            //转换完后关闭插件
            //这样做是为了解决窗口
            //再次点击时路径错误的Bug
            // window.Close();
        }

        [InfoBox("第一行字段名, 第二行中文注释，第三行类型，类型支持(int, float, bool, enum, list, dictionary)")] [ReadOnly]
        public Dictionary<string, string> example = new Dictionary<string, string>()
        {
            { "int", "1" },
            { "float", "1.0" },
            { "bool", "true" },
            { "enum|ItemTypeEnum", "A" },
            { "List|string", "JOJO,BABA" },
            { "Dictionary|string,string", "JOJO,BABA|BABA,KEKE" }
        };
        // [ShowInInspector] public string listTmp => "Type: List|string, Value: JOJO,BABA";

        public enum FormatType
        {
            JSON,
            CSV
        }

        private static IEnumerable codeTypeList = new ValueDropdownList<int>()
        {
            { "utf-8", 0 },
            { "gb2312", 1 }
        };

        private static string pathRoot;

        [MenuItem("Tools/ExcelTools")]
        private static void OpenWindow()
        {
            var window = GetWindow<ExcelToolsPlus>();
            //初始化
            pathRoot = Application.dataPath;
            //注意这里需要对路径进行处理
            //目的是去除Assets这部分字符以获取项目目录
            //我表示Windows的/符号一直没有搞懂
            pathRoot = pathRoot.Substring(0, pathRoot.LastIndexOf("/"));

            // Nifty little trick to quickly position the window in the middle of the editor.
            window.position = GUIHelper.GetEditorWindowRect().AlignCenter(700, 700);
        }

        protected override void OnGUI()
        {
            base.OnGUI();
        }

        public bool IsExcel(Object obj)
        {
            string objPath = AssetDatabase.GetAssetPath(obj);
            return objPath.EndsWith(".xlsx");
        }
    }
}