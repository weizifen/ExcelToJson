using System.Collections;
using System.Collections.Generic;
using System.Text;
using excel2json;
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
        [BoxGroup("基础配置")] [Title("资源收集")] [AssetList(Path = "Excel", CustomFilterMethod = "IsExcel")]
        public List<Object> ExcelList;

        [BoxGroup("基础配置")] [LabelText("选择格式类型")] [HideLabel] [EnumPaging]
        public ExportType exportType = ExportType.Array;

        [BoxGroup("基础配置")] [Title("选择编码类型")] [HideLabel] [ValueDropdown("codeTypeList")]
        public int codeType = 0;

        [BoxGroup("基础配置")] [ShowInInspector] [LabelText("小写")]
        public bool lowcase = false;

        [FoldoutGroup("扩展配置")] [ReadOnly] [LabelText("前三行")]
        public int headerRows = 3;

        [FoldoutGroup("扩展配置")] [LabelText("forceSheetName")]
        public bool sheetName = false;

        [FoldoutGroup("扩展配置")] [ReadOnly] [LabelText("忽略标志位")]
        public string excludePrefix = "#";

        [FoldoutGroup("扩展配置")] public bool convertJsonStringInCeil = true;

        [FoldoutGroup("扩展配置")] [Title("选择格式类型")] [HideLabel] [EnumPaging] [ReadOnly]
        public FormatType formatType = FormatType.JSON;

        [BoxGroup("基础配置")] [Title("输出文件夹(json,csv), 默认(Assets/Excel)")] [HideLabel] [FolderPath]
        public string outPath = "Assets/Excel";

        [BoxGroup("基础配置")] [Title("输出文件夹(cs), 默认(Assets/Excel)")] [HideLabel] [FolderPath]
        public string outCsPath = "Assets/Excel";

        [BoxGroup("基础配置")] [LabelText("设置命名空间")] [HideLabel]
        public string setNamespace = "HotUpdateScripts.Xiuxian";

        [FoldoutGroup("扩展配置")] [Title("保留Excel源文件")] [ShowInInspector] [LabelText("保留源文件")]
        public bool keepSource = true;


        [Button(ButtonSizes.Gigantic)]
        [LabelText("生成代码并生成Code")]
        public void GenJson()
        {
            this.Exec();
        }

        public void Exec()
        {
            foreach (Object obj in ExcelList)
            {
                var assetsPath = AssetDatabase.GetAssetPath(obj);
                string excelName = obj.name;

                //获取Excel文件的绝对路径
                string excelPath = pathRoot + "/" + assetsPath;

                //-- Header
                int header = headerRows;

                //-- Encoding
                //判断编码类型
                Encoding encoding = null;
                if (codeType == 0)
                {
                    encoding = Encoding.GetEncoding("utf-8"); // utf8-nobom
                }
                else if (codeType == 1)
                {
                    encoding = Encoding.GetEncoding("gb2312");
                }

                //-- Export path
                string exportPath;
                string output = "";
                string outputCs = "";
                if (formatType == FormatType.JSON)
                {
                    output = pathRoot + "/" + outPath + "/" + obj.name + ".json";
                    outputCs = pathRoot + "/" + outCsPath + "/" + obj.name + ".cs";

                    // excel.ConvertToJson(output, encoding, true, outputCs);
                }

                //-- Load Excel
                ExcelLoader excel = new ExcelLoader(excelPath, headerRows);

                bool tmpExportArray = exportType == ExportType.Array;

                //-- export

                //-- 生成C#定义文件

                if (outputCs.Length > 0)
                {
                    CSDefineGenerator generator = new CSDefineGenerator(excelName, excel, excludePrefix, setNamespace);
                    generator.SaveToFile(outputCs, encoding);
                }


                JsonExporter exporter = new JsonExporter(excel, lowcase, tmpExportArray, "yyyy/MM/dd", sheetName,
                    header, excludePrefix, convertJsonStringInCeil, false);
                exporter.SaveToFile(output, encoding);


                //判断是否保留源文件
                if (!keepSource)
                {
                    FileUtil.DeleteFileOrDirectory(excelPath);
                }

                //刷新本地资源
                AssetDatabase.Refresh();
            }

            //转换完后关闭插件
            //这样做是为了解决窗口
            //再次点击时路径错误的Bug
            // window.Close();
            Close();
        }

        [InfoBox("第一行字段名, 第二行中文注释，第三行类型，类型支持(int, float, bool, list, dictionary)")] [ReadOnly]
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

        public enum ExportType
        {
            Array,
            DictObject
        }

        public enum FormatType
        {
            JSON,
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