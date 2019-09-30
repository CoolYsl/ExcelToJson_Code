using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Diagnostics;

namespace ExcelToJson
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter {
        string mContext = "";
        public string FilePath;
        public Encoding encoding;
        public string context {
            get {
                return mContext;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>
        public JsonExporter(DataSet dataSet, bool lowcase, bool exportArray, string dateFormat,string filepath,Encoding en) {

            FilePath = filepath;
            encoding = en;
            List<DataTable> validSheets = new List<DataTable>();
            for (int i = 0; i < dataSet.Tables.Count; i++) {
                DataTable sheet = dataSet.Tables[i];

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheet);
            }

            var jsonSettings = new JsonSerializerSettings {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

			Stopwatch sw1 = new Stopwatch();
			Stopwatch sw2 = new Stopwatch();
			foreach (var sheet in validSheets)
            {

                Dictionary<string, object> data = new Dictionary<string, object>();
				sw1.Start();
                object sheetValue = convertSheet(sheet, exportArray, lowcase);
                data.Add(sheet.TableName, sheetValue);
                if (sheet.Rows.Count <= 2 && sheet.Columns.Count <= 2)
                    continue;
                mContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
				sw1.Stop();
				sw2.Start();
                SaveToFile(sheet.TableName);
				sw2.Stop();
            }
			TimeSpan ts1 = sw1.Elapsed;
			Console.WriteLine("处理Json数据总共花费{0}ms.", ts1.TotalMilliseconds);
			TimeSpan ts2 = sw2.Elapsed;
			Console.WriteLine("写入Json文件总共花费{0}ms.", ts2.TotalMilliseconds);


		}

        private object convertSheet(DataTable sheet, bool exportArray, bool lowcase) {
            if (exportArray)
                return convertSheetToArray(sheet, lowcase);
            else
                return convertSheetToDict(sheet, lowcase);
        }

        private object convertSheetToArray(DataTable sheet, bool lowcase) {
            List<object> values = new List<object>();

            int firstDataRow = 2;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row, lowcase, firstDataRow)
                    );
            }

            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase) {
            Dictionary<string, object> importData =
                new Dictionary<string, object>();

            int firstDataRow = 2;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = convertRowToDict(sheet, row, lowcase, firstDataRow);
                //rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row, bool lowcase, int firstDataRow) {
            var rowData = new Dictionary<string, object>();
            int col = 1;
            foreach (DataColumn column in sheet.Columns) {
                object value = row[column];

                if (value.GetType() == typeof(DBNull)) {
                    value = getColumnDefault(sheet, column, firstDataRow);
                }
                else if (value.GetType() == typeof(string)) { // 去掉数值字段的“.0”
					string str = value as string;
					int tmpInt;
                    double tmpDouble;
					if(str.Substring(str.Length-1,1).Equals("%"))
					{
						
					}
					if (int.TryParse(str, out tmpInt))
					{
						value = tmpInt;
					}
					else if (str.Substring(str.Length - 1, 1).Equals("%"))
					{
						str = str.Substring(0, str.Length - 1);
						if (Double.TryParse(str, out tmpDouble))
							value = tmpDouble*0.01;
					}
					else if (Double.TryParse(str, out tmpDouble))
						value = tmpDouble;


				}
                else if(value.GetType() == typeof(double))
                {
                    double num = (double)value;
                    if (num % 1 == 0)
                    { 
                        value = (int)num;
                    }
                }

                string fieldName = column.ToString();
                // 表头自动转换成小写
                if (lowcase)
                    fieldName = fieldName.ToLower();

                if (string.IsNullOrEmpty(fieldName))
                    fieldName = string.Format("col_{0}", col);

                rowData[fieldName] = value;
                col++;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, DataColumn column, int firstDataRow) {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++) {
                object value = sheet.Rows[i][column];
                Type valueType = value.GetType();
                if (valueType != typeof(System.DBNull)) {
                    if (valueType.IsValueType)
                        return Activator.CreateInstance(valueType);
                    break;
                }
            }
            return "";
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string fileName) {
            //-- 保存文件
            using (FileStream file = new FileStream(FilePath + @"\" + fileName + ".json", FileMode.Create, FileAccess.Write)) {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(mContext);
            }
        }
    }
}
