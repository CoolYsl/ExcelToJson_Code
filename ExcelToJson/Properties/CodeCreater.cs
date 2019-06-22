using System.Collections.Generic;
using System;
using System.IO;
using System.Data;
using System.Text;

namespace ExcelToJson.Properties
{
    public enum CreateType
    {
        CPP = 1,
        CSharp = 2
    }

    //代码生成器
    public abstract class ACodeCreater
    {
        public string FilePath;
        public string FileName;
        public StringBuilder sb = new StringBuilder();
        public string Suffix;
        public ACodeCreater(string filePath, string fileName)
        {
            FilePath = filePath;
            FileName = fileName;
        }
        public const string _namespace = "JsonReadObject";
        public abstract void CodeCreat(List<DataTable> validSheets);
        public virtual void SaveToFile(Encoding encoding)
        {
            try
            {
                
                if(!Directory.Exists(FilePath))
                {
                    Console.WriteLine(FilePath);
                    Directory.CreateDirectory(FilePath);
                    
                }
                using (FileStream file = new FileStream(FilePath +@"\"+ FileName + Suffix, FileMode.Create, FileAccess.Write))
                {
                    Console.WriteLine(FilePath + FileName + Suffix);
                    using (TextWriter writer = new StreamWriter(file, encoding))
                        writer.Write(sb.ToString());
                }
                sb.Clear();
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }

    class CSharpCodeCreater : ACodeCreater
    {

        public CSharpCodeCreater(string filePath ,string fileName):base(filePath,fileName)
        {
            Suffix = ".cs";
        }
        public override void CodeCreat(List<DataTable> validSheets)
        {
            try
            {
                sb.Clear();
				sb.AppendLine("using System; \n");
				sb.AppendLine("namespace " + ACodeCreater._namespace);
                sb.AppendLine("{");

                //遍历Sheet
                foreach (var sheet in validSheets)
                {
                    sb.AppendLine("public class " + sheet.TableName);
                    sb.AppendLine("{");
                    int col = 0;
                    //查找字段类型和字段名
                    foreach (DataColumn column in sheet.Columns)
                    {
                        string fieldName = column.ToString();
                        string Notes = sheet.Rows[1][column].ToString();
                        if (!string.IsNullOrEmpty(Notes))
						{
							var newStrs = Notes.Split('\n');
							foreach (var n in newStrs)
								sb.AppendLine("//"+ n);
						}
                            
                        if (string.IsNullOrEmpty(fieldName))
                        {
                            fieldName = string.Format("col_{0}", col);
                            col++;
                        }
                        string fieldType = sheet.Rows[0][column].ToString();
                        if (string.IsNullOrEmpty(fieldType))
                        {
                            fieldType = "string";
                        }
                        //字符转换为小写
                        fieldType = fieldType.ToLower();
                        sb.AppendLine("public "+ fieldType + " " + fieldName + ";");
                    }
                    sb.AppendLine("}\n");
                }
                sb.AppendLine("}");
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        
    }
    class CPPCodeCreater :ACodeCreater
    {
        public CPPCodeCreater(string filePath, string fileName):base(filePath, fileName)
        {
            Suffix = ".h";
        }
        public override void CodeCreat(List<DataTable> validSheets)
        {
            try
            {
                sb.Clear();
                sb.AppendLine("namespace " + ACodeCreater._namespace);
                sb.AppendLine("{");
                
                //遍历Sheet
                foreach (var sheet in validSheets)
                {
                    sb.AppendLine("class " + sheet.TableName);
                    sb.AppendLine("{");
                    sb.AppendLine("public:");
                    int col = 0;
                    //查找字段类型和字段名
                    foreach (DataColumn column in sheet.Columns)
                    {
                        string fieldName = column.ToString();
                        string Notes = sheet.Rows[1][column].ToString();
                        if (!string.IsNullOrEmpty(Notes))
						{
							var newStrs = Notes.Split('\n');
							foreach (var n in newStrs)
								sb.AppendLine("//" + n);
						}
						if (string.IsNullOrEmpty(fieldName))
                        {
                            fieldName = string.Format("col_{0}", col);
                            col++;
                        }
                        string fieldType = sheet.Rows[0][column].ToString();
                        if (string.IsNullOrEmpty(fieldType))
                        {
                            fieldType ="string";
                        }
                        //字符转换为小写
                        fieldType = fieldType.ToLower();
                        sb.AppendLine(fieldType + " " + fieldName + ";");
                    }
                    sb.AppendLine("};\n");
                }
                sb.AppendLine("}");
            }
            catch
            {
                throw new Exception("CSharpCodeCreater.CodeCreat Default!");
            }
        }
       
    }
    public class CodeCreaterManager
    {
        public List<ACodeCreater> CodeCreater;
        private string FilePath;
        private Encoding encoding;
        public CodeCreaterManager(string filepath, Encoding en)
        {
            FilePath = filepath;
            encoding = en;
            CodeCreater = new List<ACodeCreater>();
            
        }
        public void AddCreatCodeType( CreateType type)
        {
            switch (type)
            {
                case CreateType.CPP:
                    CodeCreater.Add(new CPPCodeCreater(FilePath,ACodeCreater._namespace));
                    break;
                case CreateType.CSharp:
                    CodeCreater.Add(new CSharpCodeCreater(FilePath, ACodeCreater._namespace));
                    break;
                default:
                    break;
            }
        }
        public void CodeCreat(DataSet dataSet)
        {
            try
            {
                List<DataTable> validSheets = new List<DataTable>();
                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    DataTable sheet = dataSet.Tables[i];

                    if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                        validSheets.Add(sheet);
                }
                foreach (var item in CodeCreater)
                {
                    item.CodeCreat(validSheets);
                    item.SaveToFile(encoding);
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
