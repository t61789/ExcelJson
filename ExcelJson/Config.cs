using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelJson;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Util;

namespace ExcelProtobuf
{
    public class Config
    {
        public static Config Instance;

        public string ConfigPath;
        public string ExcelDirectory;
        public string LogPath;

        public XDocument ConfigDoc;

        private XElement _Root;
        private XElement _MappingNode;

        static Config()
        {
            Instance = new Config
            {
                ConfigPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"config.xml",
                ExcelDirectory = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"excel" + Path.DirectorySeparatorChar,
                LogPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase + @"log.log",
            };

            if (!Directory.Exists(Instance.ExcelDirectory)) Directory.CreateDirectory(Instance.ExcelDirectory);
            if (!File.Exists(Instance.LogPath)) File.Create(Instance.LogPath).Close();

            Instance.LoadConfig();
        }

        public void LoadConfig()
        {
            ConfigDoc = XDocument.Load(ConfigPath);
            _Root = ConfigDoc.Root;
            _MappingNode = _Root.Element("mapping");
        }

        public void SaveConfig()
        {
            ConfigDoc.Save(ConfigPath);
        }

        public string GetMapping(string json)
        {
            return _MappingNode.Elements().FirstOrDefault(x => x.Element("json").Value == json)?.Element("excel").Value;
        }

        public IEnumerable<(string jsonPath, string excelName)> GetMappings()
        {
            return from map in _MappingNode.Elements()
                let result = (map.Element("json").Value, map.Element("excel").Value)
                select result;
        }

        public bool ExcelExists(string excel)
        {
            return _MappingNode.Elements().FirstOrDefault(x => x.Element("excel").Value == excel) != null;
        }

        public static string GetHash(string filePath)
        {
            string sha256;
            SHA256Managed s = new SHA256Managed();
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                sha256 = BitConverter.ToString(s.ComputeHash(fs));
            }
            return sha256;
        }

        public bool AddNewMappingAndExcel(string jsonPath, bool cover)
        {
            XElement tempE = _MappingNode.Elements().FirstOrDefault(x => x.Element("json").Value == jsonPath);
            if (tempE != null)  // 检查映射重复
            {
                if (cover)
                {
                    File.Delete(jsonPath);
                    File.Delete(ExcelDirectory + tempE.Element("excel").Value);
                    tempE.Remove();
                }
                else
                {
                    Program.Log($"已存在 {jsonPath} 的映射");
                    return false;
                }
            }

            string excelName = Path.GetFileName(jsonPath);
            int lastIndex = excelName.LastIndexOf('.');
            if (lastIndex == -1)
                excelName += ".xlsx";
            else
                excelName = excelName.Substring(0, lastIndex) + ".xlsx";

            string curdir = excelName;
            while (ExcelExists(excelName))  // 不断细化excel文件名直到不重名
            {
                curdir = Path.GetDirectoryName(curdir);
                if (curdir == "")
                {
                    Program.Log("Excel缓存重名");
                    return false;
                }
                string temp = Path.GetFileName(curdir);
                if (temp == "")
                    temp = curdir[0].ToString();

                excelName = temp + '_' + excelName;
            }

            using (FileStream fs = File.Create(jsonPath))
            {
                new StreamWriter(fs).Write("{\"Fields\":[],\"Rows\":[]}");
            }

            CreateEmptyExcel(excelName);

            XElement datE = new XElement("json") { Value = jsonPath };
            XElement excelE = new XElement("excel") { Value = excelName };
            XElement map = new XElement("map");
            map.Add(datE);
            map.Add(excelE);
            _MappingNode.Add(map);

            SaveConfig();

            return true;
        }

        private void CreateEmptyExcel(string excelName)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            ICellStyle style = workbook.CreateCellStyle();
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = 42;
            style.BorderLeft =
            style.BorderTop = 
            style.BorderRight = 
            style.BorderBottom = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            sheet.CreateRow(0).RowStyle = style;
            sheet.CreateRow(1).RowStyle = style;

            XSSFDataValidationHelper helper = new XSSFDataValidationHelper((XSSFSheet)sheet);
            XSSFDataValidationConstraint constraint =
                (XSSFDataValidationConstraint)helper.CreateExplicitListConstraint(Enum.GetNames(typeof(DataConverter.DataType)));
            CellRangeAddressList range = new CellRangeAddressList(1, 1, 0, 255);
            sheet.AddValidationData((XSSFDataValidation)helper.CreateValidation(constraint, range));

            style = workbook.CreateCellStyle();
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = 42;
            style.BorderLeft = 
            style.BorderTop =
            style.BorderRight = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Medium;
            style.Alignment = HorizontalAlignment.Center;
            sheet.CreateRow(2).RowStyle = style;

            using (FileStream fs = new FileStream(ExcelDirectory + excelName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
            workbook.Close();

        }

        public string[] GetEmptyMapping()
        {
            return ConfigDoc.Root.Element("mapping").Elements().
                Select(x => x.Element("json").Value).
                Where(x => !File.Exists(x)).ToArray();
        }

        public void DeleteEmptyMapping(string[] emptyMappings)
        {
            var result = (from i in _MappingNode.Elements()
                          let temp = i.Element("json").Value
                          where emptyMappings.Contains(temp)
                          select i).ToArray();
            foreach (var item in result)
            {
                try
                {
                    File.Delete(ExcelDirectory + item.Element("excel").Value);
                    Program.Log($"删除成功 {item.Element("json").Value}");
                    item.Remove();
                }
                catch (Exception)
                {
                    Program.Log($"删除失败 {item.Element("json").Value}");
                }
            }

            SaveConfig();
        }

        public string[] GetInvalidExcel()
        {
            return (from item in Directory.GetFiles(ExcelDirectory, "*.xlsx")
                    let excel = Path.GetFileName(item)
                    where !ExcelExists(excel)
                    select item).ToArray();
        }

        public void DeleteInvalidExcel(string[] excelPaths)
        {
            foreach (var item in excelPaths)
            {
                try
                {
                    File.Delete(item);
                    Program.Log($"删除成功 {item}");
                }
                catch (Exception)
                {
                    Program.Log($"删除失败 {item}");
                }
            }
        }

        public void RecordLog(Exception e)
        {
            File.AppendAllText(LogPath, $"[{DateTime.Now}] :>{e}\n");
        }

        public bool CheckExcelHash(string excelName,bool setNew)
        {
            XElement e = _MappingNode.Elements().FirstOrDefault(x => x.Element("excel").Value == excelName);
            string newHash = GetHash(ExcelDirectory + excelName);
            bool result = e?.Attribute("hash")?.Value == newHash;
            if (!setNew|| result || e == null) return result;
            e.SetAttributeValue("hash",newHash);
            SaveConfig();
            return false;
        }
    }
}
