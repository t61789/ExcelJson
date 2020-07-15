using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using ExcelJson;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

namespace ExcelProtobuf
{
    public class DataConverter
    {
        public static void Process(bool force)
        {
            Program.Log("***开始转换数据***");

            foreach (var (jsonPath, excelName) in Config.Instance.GetMappings())
            {
                try
                {
                    if(!Config.Instance.CheckExcelHash(excelName,false) || force || !File.Exists(jsonPath))
                        ProcessData(jsonPath, excelName);
                }
                catch (Exception e)
                {
                    Config.Instance.RecordLog(e);
                    Program.Log($"数据转换失败 {jsonPath}");
                    continue;
                }
                Program.Log($"数据转换成功 {jsonPath}");
            }

            Program.Log("***数据转换完成***");
        }

        private static readonly StringBuilder JsonBuilder = new StringBuilder();

        private static void ProcessData(string jsonPath, string excelName)
        {
            XSSFWorkbook workbook;
            using (FileStream fs = new FileStream(Config.Instance.ExcelDirectory + excelName, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
            }
            ISheet sheet = workbook.GetSheetAt(0);

            JsonBuilder.Clear();
            JsonBuilder.Append("{\"Fields\":[");

            int fieldsCount = 0;
            IRow dataNameRow = sheet.GetRow(2);
            if(dataNameRow!=null)
                foreach (var cell in dataNameRow)
                {
                    if(fieldsCount!=0)
                        JsonBuilder.Append(',');
                    JsonBuilder.Append($"\"{cell}\"");
                    fieldsCount++;
                }
            JsonBuilder.Append("],\"Rows\":[");

            var dataType = new DataType[fieldsCount];
            IRow dataTypeRow = sheet.GetRow(1);
            for (int i = 0; i < fieldsCount; i++)
                dataType[i] = (DataType)Enum.Parse(typeof(DataType), dataTypeRow.GetCell(i).ToString());

            foreach (IRow row in sheet)
            {
                if (row.RowNum < 3) continue;
                JsonBuilder.Append('[');
                for (int i = 0; i < fieldsCount; i++)
                {
                    JsonBuilder.Append(FormatData(row.GetCell(i)?.ToString(), dataType[i]));
                    if(i!=fieldsCount-1)
                        JsonBuilder.Append(',');
                }
                JsonBuilder.Append(']');
                if (row.RowNum != sheet.LastRowNum)
                    JsonBuilder.Append(',');
            }
            JsonBuilder.Append("]}");

            using (FileStream fs = new FileStream(jsonPath, FileMode.Create, FileAccess.Write))
            {
                var bytes = Encoding.UTF8.GetBytes(JsonBuilder.ToString());
                
                fs.Write(bytes, 0, bytes.Length);
            }

            Config.Instance.CheckExcelHash(excelName, true);
        }

        private static string FormatData(string data,DataType type)
        {
            switch (type)
            {
                case DataType.Integer:
                    data = string.IsNullOrEmpty(data) ? "0" : data;
                    return int.Parse(data).ToString();
                case DataType.Float:
                    data = string.IsNullOrEmpty(data) ? "0" : data;
                    data = float.Parse(data).ToString(CultureInfo.CurrentCulture);
                    return data.IndexOf('.') == -1 ? data + ".0" : data;
                case DataType.String:
                    return $"\"{data}\"";
                case DataType.Bool:
                    data = string.IsNullOrEmpty(data) ? "false" : data;
                    return bool.Parse(data).ToString();
                case DataType.Unknown:
                    throw new InvalidCastException();
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }
        }

        public enum DataType
        {
            Unknown = 0,
            Integer = 1,
            Float = 2,
            String = 3,
            Bool = 4
        }
    }
}