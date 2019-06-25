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
        public StringBuilder sbObj = new StringBuilder();
		public List<string> SheetNames = new List<string>();
		public string Suffix;
        public ACodeCreater(string filePath, string fileName)
        {
            FilePath = filePath;
            FileName = fileName;
        }
        public const string _namespace = "JsonReadObject";
        public abstract void ObjCodeCreat(List<DataTable> validSheets);
		public abstract void JsonReaderCodeCreat(List<DataTable> validSheets);
        public virtual void SaveObjCodeToFile(Encoding encoding)
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
                        writer.Write(sbObj.ToString());
                }
                sbObj.Clear();
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
		public void InitSheetName(List<DataTable> validSheets)
		{ 
			foreach (var sheet in validSheets)
			{
				SheetNames.Add(sheet.TableName);
			}
		}

    }

    class CSharpCodeCreater : ACodeCreater
    {

        public CSharpCodeCreater(string filePath ,string fileName):base(filePath,fileName)
        {
            Suffix = ".cs";
        }

		public override void JsonReaderCodeCreat(List<DataTable> validSheets)
		{
		}

		public override void ObjCodeCreat(List<DataTable> validSheets)
        {
            try
            {
                sbObj.Clear();
				sbObj.AppendLine("using System; \n");
				sbObj.AppendLine("namespace " + ACodeCreater._namespace);
                sbObj.AppendLine("{");

                //遍历Sheet
                foreach (var sheet in validSheets)
                {
                    sbObj.AppendLine("public class " + sheet.TableName);
                    sbObj.AppendLine("{");
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
								sbObj.AppendLine("//"+ n);
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
                        sbObj.AppendLine("public "+ fieldType + " " + fieldName + ";");
                    }
                    sbObj.AppendLine("}\n");
                }
                sbObj.AppendLine("}");
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        
    }
    class CPPCodeCreater :ACodeCreater
    {
		private StringBuilder hsb = new StringBuilder();
		private StringBuilder cppsb = new StringBuilder();
		public CPPCodeCreater(string filePath, string fileName):base(filePath, fileName)
        {
            Suffix = ".h";
        }
		private void SaveJsonReaderToFile()
		{
			Encoding utf8 = new UTF8Encoding(false);
			try
			{
				if (!Directory.Exists(FilePath))
				{
					Directory.CreateDirectory(FilePath);
				}
				using (FileStream file = new FileStream(FilePath + @"\JsonReader"  + Suffix, FileMode.Create, FileAccess.Write))
				{
					using (TextWriter writer = new StreamWriter(file, utf8))
						writer.Write(hsb.ToString());
				}
				using (FileStream file = new FileStream(FilePath + @"\JsonReader.cpp", FileMode.Create, FileAccess.Write))
				{
					using (TextWriter writer = new StreamWriter(file, utf8))
						writer.Write(cppsb.ToString());
				}
				hsb.Clear();
				cppsb.Clear();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
		private void CreatheaderFile()
		{
			hsb.Clear();
			hsb.AppendLine("#pragma once ");
			hsb.AppendLine("#include \"JsonReadObject.h\"");
			hsb.AppendLine("#include \"cJson/CJsonObject.hpp\"");
			hsb.AppendLine("# include <vector>");
			hsb.AppendLine("# include <fstream>");
			hsb.AppendLine("using namespace JsonReadObject;");
			hsb.AppendLine("class JsonReader");
			hsb.AppendLine("{");
			hsb.AppendLine("public:");
			hsb.AppendLine("\tJsonReader();");
			hsb.AppendLine("\t~JsonReader();");
			hsb.AppendLine("\tvoid LoadAll();");
			hsb.AppendLine("public:");
			foreach (var name in SheetNames)
			{
				hsb.AppendLine("\tstd::vector <" + name + "> Get" + name + "();");
			}
			hsb.AppendLine();
			hsb.AppendLine("private:");
			foreach(var name in SheetNames)
			{
				hsb.AppendLine("\tvoid Load"+name+"();");
			}
			hsb.AppendLine("\n");
			hsb.AppendLine("private:");
			hsb.AppendLine("\tchar* GetFileStr(const char* fileName);\n");
			hsb.AppendLine("private:");
			hsb.AppendLine("\tneb::CJsonObject cJsonObj;");
			hsb.AppendLine("\tstatic ifstream inFile;");
			hsb.AppendLine("\tchar buffer[2048];");
			hsb.AppendLine("private:");
			foreach (var name in SheetNames)
			{
				hsb.AppendLine("\tstd::vector <" + name + "> vec"+name+";");
			}
			hsb.AppendLine("};");
		}
		private List<string> GetFixedName(List<DataTable> validSheets,string name)
		{
			List<string> str = new List<string>();
			foreach (var sheet in validSheets)
			{
				if (name.Equals(sheet.TableName))
				{
					int col = 0;
					foreach (DataColumn column in sheet.Columns)
					{
						string fieldName = column.ToString();
						if (string.IsNullOrEmpty(fieldName))
						{
							fieldName = string.Format("col_{0}", col);
							col++;
						}
						str.Add(fieldName);
					}
					return str;
				}
			}
			return str;
		}
		private void CreatCppFile(List<DataTable> validSheets)
		{
			cppsb.Clear();
			cppsb.AppendLine("#include \"JsonReader.h\"");
			cppsb.AppendLine("#include \"cJson/CJsonObject.hpp\"");
			cppsb.AppendLine("# include \"define.h\"");
			cppsb.AppendLine("#include <fstream>");
			cppsb.AppendLine("");
			cppsb.AppendLine("ifstream JsonReader::inFile;");
			cppsb.AppendLine("const int bufferLen = 2048;");
			cppsb.AppendLine("JsonReader::JsonReader():buffer{0}");
			cppsb.AppendLine("{");
			cppsb.AppendLine("");
			cppsb.AppendLine("}");
			cppsb.AppendLine("");
			cppsb.AppendLine("JsonReader::~JsonReader()");
			cppsb.AppendLine("{");
			cppsb.AppendLine("");
			cppsb.AppendLine("}");
			cppsb.AppendLine("");
			cppsb.AppendLine("void JsonReader::LoadAll()");
			cppsb.AppendLine("{");
			foreach(var name in SheetNames)
			{
				cppsb.AppendLine("\tLoad" + name + "();");
			}
			cppsb.AppendLine("}");
			cppsb.AppendLine("");
			foreach (var name in SheetNames)
			{
				cppsb.AppendLine("std::vector <" + name + "> JsonReader::Get" + name + "()");
				cppsb.AppendLine("{");
				cppsb.AppendLine("\treturn vec" + name + ";");
				cppsb.AppendLine("}\n");
			}
			foreach (var name in SheetNames)
			{
				cppsb.AppendLine("void JsonReader::Load" + name + "()");
				cppsb.AppendLine("{");
				cppsb.AppendLine("\tOUR_DEBUG((LM_INFO,\"Load" + name + ".json....\"));");
				cppsb.AppendLine("");
				cppsb.AppendLine("\tGetFileStr(\"./ Conf / " + name + ".json\");");
				cppsb.AppendLine("\tcJsonObj.Clear();");
				cppsb.AppendLine("\tcJsonObj.Parse(buffer);");
				cppsb.AppendLine("\tint arrlen = cJsonObj.GetArraySize();");
				cppsb.AppendLine("\tfor (int i=0;i<arrlen;++i)");
				cppsb.AppendLine("\t{");
				cppsb.AppendLine("\t\t" + name + " conf;");

				foreach (var fixedName in GetFixedName(validSheets, name))
				{
					cppsb.AppendLine("\t\tcJsonObj[i].Get(\"" + fixedName + "\" , conf." + fixedName + ");");
				}
				cppsb.AppendLine("\t\tvec" + name + ".push_back(conf);");
				cppsb.AppendLine("\t}");
				cppsb.AppendLine("}");
				cppsb.AppendLine("");
			}
			cppsb.AppendLine("char* JsonReader::GetFileStr(const char* fileName)");
			cppsb.AppendLine("{");
			cppsb.AppendLine("\tinFile.open(fileName, std::ios::in);");
			cppsb.AppendLine("\tinFile.seekg(0, std::ios::end);");
			cppsb.AppendLine("\tint len = inFile.tellg();");
			cppsb.AppendLine("\tif (len < bufferLen)");
			cppsb.AppendLine("\t{");
			cppsb.AppendLine("\t\tinFile.seekg(0, std::ios::beg);");
			cppsb.AppendLine("\t\tinFile.read(buffer, len);");
			cppsb.AppendLine("\t}");
			cppsb.AppendLine("\telse");
			cppsb.AppendLine("\t{");
			cppsb.AppendLine("\t\tOUR_DEBUG((LM_INFO, \"Buffer is verflow!!!!!!!!!!%s\", fileName));");
			cppsb.AppendLine("\t}");
			cppsb.AppendLine("\tinFile.close();");
			cppsb.AppendLine("\treturn buffer;");
			cppsb.AppendLine("}");
		}
		public override void JsonReaderCodeCreat(List<DataTable> validSheets)
		{
			CreatheaderFile();
			CreatCppFile(validSheets);
			SaveJsonReaderToFile();
		}

		public override void ObjCodeCreat(List<DataTable> validSheets)
        {
            try
            {
                sbObj.Clear();
                sbObj.AppendLine("namespace " + ACodeCreater._namespace);
                sbObj.AppendLine("{");
                
                //遍历Sheet
                foreach (var sheet in validSheets)
                {
                    sbObj.AppendLine("class " + sheet.TableName);
                    sbObj.AppendLine("{");
                    sbObj.AppendLine("public:");
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
								sbObj.AppendLine("//" + n);
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
                        sbObj.AppendLine(fieldType + " " + fieldName + ";");
                    }
                    sbObj.AppendLine("};\n");
                }
                sbObj.AppendLine("}");
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
					{
						validSheets.Add(sheet);
					}

                }
                foreach (var item in CodeCreater)
                {
					item.InitSheetName(validSheets);
                    item.ObjCodeCreat(validSheets);
					item.JsonReaderCodeCreat(validSheets);
                    item.SaveObjCodeToFile(encoding);
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
