using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ExcelProtobuf;
using Newtonsoft.Json;

namespace ExcelJson
{
    public class Program
    {
        private static readonly Process Process = new Process();

        [STAThread]
        public static void Main(string[] args)
        {
            new Program().Start();
        }

        public void Start()
        {
            while (true)
            {
                int command = Command("开始使用", "打开表文件", "添加新的映射", "更新所有数据", "强制更新所有数据", "删除闲置映射和缓存", "打开配置文件","打开日志文件", "退出");
                Config.Instance.LoadConfig();
                switch (command)
                {
                    case 1:
                        OpenFile();
                        break;
                    case 2:
                        AddNewFile();
                        break;
                    case 3:
                        DataConverter.Process(false);
                        break;
                    case 4:
                        DataConverter.Process(true);
                        break;
                    case 5:
                        DeleteEmptyMapping();
                        break;
                    case 6:
                        OpenConfigFile();
                        break;
                    case 7:
                        OpenFile(Config.Instance.LogPath);
                        break;
                    case 8:
                        Environment.Exit(0);
                        break;
                }
            }
        }

        public void OpenFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "json文件|*.bytes|dat数据库|*.dat|所有文件|*.*",
                Title = "保存表文件"
            };

            DialogResult d = openFileDialog.ShowDialog();
            if (d != DialogResult.OK) return;
            string excelName = Config.Instance.GetMapping(openFileDialog.FileName);
            if (excelName == null)
            {
                switch (Command($"到 {openFileDialog.FileName} 的映射不存在，是否创建并覆盖文件？","是","否"))
                {
                    case 1:
                        if (Config.Instance.AddNewMappingAndExcel(openFileDialog.FileName, true))
                        {
                            Log($"映射创建成功 {openFileDialog.FileName}");
                            excelName = Config.Instance.GetMapping(openFileDialog.FileName);
                            break;
                        }
                        else
                        {
                            Log($"映射创建失败 {openFileDialog.FileName}");
                            return;
                        }
                    case 2:
                        return;
                }
            }

            OpenFile(Config.Instance.ExcelDirectory + excelName);
        }

        public void AddNewFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "json文件|*.bytes|dat数据库|*.dat|所有文件|*.*",
                Title = "保存表文件",
            };
            DialogResult d = saveFileDialog.ShowDialog();

            if (d != DialogResult.OK) return;

            string jsonPath = saveFileDialog.FileName;
            Log(Config.Instance.AddNewMappingAndExcel(jsonPath, true)
                ? $"映射创建成功 {jsonPath}"
                : $"映射创建失败 {jsonPath}");

            if (Command("是否打开文件？","是","否")==1)
                OpenFile(Config.Instance.ExcelDirectory + Config.Instance.GetMapping(jsonPath));
        }

        public void DeleteEmptyMapping()
        {
            Log("开始扫描闲置映射……");

            var del = Config.Instance.GetEmptyMapping();
            if (del.Length == 0)
                Log("无闲置映射");
            else
            {
                foreach (var item in del)
                    Console.WriteLine('\n' + item);
                if (Command("是否删除以上闲置映射及其Excel缓存?", "是", "否") == 1)
                    Config.Instance.DeleteEmptyMapping(del);
            }

            Log("开始扫描闲置Excel缓存……");
            del = Config.Instance.GetInvalidExcel();
            if (del.Length == 0)
                Log("无闲置Excel缓存");
            else
            {
                foreach (var item in del)
                    Console.WriteLine('\n'+Path.GetFileName(item));
                if (Command("是否删除以上闲置Excel缓存?", "是", "否") == 1)
                    Config.Instance.DeleteInvalidExcel(del);
            }
        }

        public void OpenConfigFile()
        {
            OpenFile(Config.Instance.ConfigPath);
        }

        public static void OpenFile(string path)
        {
            Process.StartInfo = new ProcessStartInfo()
            {
                FileName = path
            };
            Process.Start();
            Process.Close();
        }

        public static void Exec(string file, string arg)
        {
            Process.StartInfo = new ProcessStartInfo()
            {
                FileName = file,
                Arguments = arg,
                UseShellExecute = false,
                RedirectStandardOutput = true
            };
            Process.Start();
            Process.WaitForExit();
            Process.StandardOutput.ReadToEnd();
            Process.Close();
        }

        public static void Log(string message, params object[] content)
        {
            Console.WriteLine(DateTime.Now + " :>" + message, content);
        }
        
        public static int Command(string discripe, params string[] selection)
        {
            while (true)
            {
                Console.WriteLine("\n[" + discripe + "]");
                for (int i = 1; i <= selection.Length; i++)
                    Console.WriteLine(i.ToString() + '.' + selection[i - 1]);
                Console.Write("command:>");
                if (int.TryParse(Console.ReadLine(), out int result))
                {
                    if (result >= 1 && result <= selection.Length)
                        return result;
                }
                Console.WriteLine();
            }
        }
    }
}
