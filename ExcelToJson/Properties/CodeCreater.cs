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
		public abstract void SkillCodeCreat(List<DataTable> validSheets);
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

		public override void SkillCodeCreat(List<DataTable> validSheets)
		{
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
		private void SaveToFile(StringBuilder sb, string dic,string filename)
		{
			Encoding utf8 = new UTF8Encoding(false);
			try
			{
				if (!Directory.Exists(FilePath+@"\"+dic))
				{
					Directory.CreateDirectory(FilePath + @"\" + dic);
				}
				using (FileStream file = new FileStream(FilePath+@"\"+dic + @filename, FileMode.Create, FileAccess.Write))
				{
					using (TextWriter writer = new StreamWriter(file, utf8))
						writer.Write(sb.ToString());
				}
				
				sb.Clear();
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
				hsb.AppendLine("\tconst std::vector <" + name + ">& Get" + name + "()const;");
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
				cppsb.AppendLine("const std::vector <" + name + ">& JsonReader::Get" + name + "()const");
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
			SaveToFile(hsb, "",@"\JsonReader.h");
			SaveToFile(cppsb,"", @"\JsonReader.cpp");
		}

		public override void ObjCodeCreat(List<DataTable> validSheets)
        {
            try
            {
                sbObj.Clear();
				sbObj.AppendLine("#include \"DataFormat.h\"");
				sbObj.AppendLine("typedef uint32 uint;");
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

		public override void SkillCodeCreat(List<DataTable> validSheets)
		{
			try
			{

			
			//查询技能表
			List<DataTable> skillsconf = new List<DataTable>();
			foreach (var sheet in validSheets)
			{
				if(sheet.TableName.Contains("Skill_"))
				{
					skillsconf.Add(sheet);
				}
			}
			if(skillsconf.Count == 0)
			{
				return;
			}

			//判断重复字段
			Dictionary<string, Dictionary<string, string>> fieldAndType = new Dictionary<string, Dictionary<string, string>>();
			Dictionary<string, int> fieldNameCount = new Dictionary<string, int>();
			foreach (var sheet in skillsconf)
			{
				int col = 0;
				string skillname = sheet.TableName.Replace("Conf", "");
				fieldAndType[skillname] = new Dictionary<string, string>();
				foreach (DataColumn column in sheet.Columns)
				{
					//字段名
					string fieldName = column.ToString();
					if (string.IsNullOrEmpty(fieldName))
					{
						fieldName = string.Format("col_{0}", col);
						col++;
					}
					//字段类型
					DataRow row = sheet.Rows[0];
					string fieldType = row[column].ToString();
					if (!fieldNameCount.ContainsKey(fieldName))
					{
						fieldNameCount[fieldName] = 1;
					}
					else
					{
						fieldNameCount[fieldName] += 1;
					}

					fieldAndType[sheet.TableName.Replace("Conf","")][fieldName] = fieldType;
				}

			}
			Dictionary<string, string> repeatfieldType = new Dictionary<string, string>();
			//重复的字段找出来写入基类
			foreach (var rep in fieldNameCount)
			{
				if (rep.Value == skillsconf.Count)
				{
					repeatfieldType[rep.Key] = fieldAndType[skillsconf[0].TableName.Replace("Conf","")][rep.Key];
					foreach(var skill in fieldAndType)
					{
						skill.Value.Remove(rep.Key);
					}
				}
			}
				CreatSkillBaseClass(repeatfieldType);
				CreatSkillClass(fieldAndType);
			}
			catch
			{
				throw (new Exception("读取数据错误"));
			}

		}
		#region//创建技能类
		private void CreatSkillClass(Dictionary<string, Dictionary<string, string>> fieldAndType)
		{
			StringBuilder sb = new StringBuilder();
			foreach (var skill in fieldAndType)
			{
				CreatSkillHeaderFile(skill.Key, skill.Value, sb);
				SaveToFile(sb, "CPP", @"\" + skill.Key + ".h");
				CreatSkillCppFile(skill.Key, skill.Value, sb);
				SaveToFile(sb, "CPP", @"\" + skill.Key + ".cpp");
			}
		}
		private void CreatSkillHeaderFile(string skillName, Dictionary<string, string> fieldAndType, StringBuilder sb)
		{
			sb.Clear();
			sb.AppendLine("#pragma once");
			sb.AppendLine("#include \"BaseSkill.h\"");
			sb.AppendLine("class " + skillName + ":public BaseSkill");
			sb.AppendLine("{");
			sb.AppendLine("public:");
			sb.AppendLine("\t" + skillName + "();");
			sb.AppendLine("\tvirtual ~" + skillName + "();");
			sb.AppendLine("\tvirtual bool Release();");
			sb.AppendLine("\tvirtual bool Init(int level);\n");
			sb.AppendLine("public:");

			foreach (var field in fieldAndType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("\tconst string& Get" + field.Key + "()const;");
				}
				else
				{
					sb.AppendLine("\t" + field.Value + " Get" + field.Key + "()const;");
				}
			}
			sb.AppendLine("\nprotected:");
			foreach (var field in fieldAndType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("\tvoid Set" + field.Key + "(const string& value);");
				}
				else
				{
					sb.AppendLine("\tvoid Set" + field.Key + "(" + field.Value + " value);");
				}
			}
			sb.AppendLine("\nprivate:");
			foreach (var field in fieldAndType)
			{
				sb.AppendLine("\t" + field.Value + " " + field.Key + ";");
			}
			sb.AppendLine("};");
		}
		private void CreatSkillCppFile(string skillName, Dictionary<string, string> fieldAndType, StringBuilder sb)
		{
			sb.Clear();
			sb.AppendLine("#include \""+ skillName + ".h\"");
			sb.AppendLine("#include \"DaMenPai.cpp\"\n");
			sb.AppendLine(skillName+"::"+ skillName+"()");
			sb.AppendLine("{\n");
			sb.AppendLine("}\n");
			sb.AppendLine(skillName + "::~" + skillName + "()");
			sb.AppendLine("{\n");
			sb.AppendLine("}\n");
			sb.AppendLine("bool "+skillName + "::Release()");
			sb.AppendLine("{\n");
			sb.AppendLine("}\n");
			sb.AppendLine("bool "+ skillName+"::Init(int level)");
			sb.AppendLine("{\n");
			sb.AppendLine("\tauto confVec = g_pJsonReader->Get" + skillName + "Conf();\n");
			sb.AppendLine("\tfor (auto it : confVec)");
			sb.AppendLine("\t{");
			sb.AppendLine("\t\tif (it.Level)");
			sb.AppendLine("\t\t{");
			foreach (var field in fieldAndType)
			{
				
				sb.AppendLine("\t\t\tSet" + field.Key + "(it." + field.Key + ");");
				
			}
			sb.AppendLine("\t\t\treturn true;");
			sb.AppendLine("\t\t}");
			sb.AppendLine("\t}");
			sb.AppendLine("\treturn false;");
			sb.AppendLine("}\n");
			foreach (var field in fieldAndType)
			{
				sb.AppendLine(field.Value + " " + skillName + "::Get" + field.Key + "()const");
				sb.AppendLine("{");
				sb.AppendLine("\treturn " + field.Key + ";");
				sb.AppendLine("}");
			}
			foreach (var field in fieldAndType)
			{
				sb.AppendLine("void " + skillName + "::Set" + field.Key + "(" + field.Value + " value)");
				sb.AppendLine("{");
				sb.AppendLine("\t" + field.Key + "= value;");
				sb.AppendLine("}");
			}
		}
		#endregion
		#region//创建技能基类
		private void CreatSkillBaseClass(Dictionary<string, string> repeatfieldType)
		{
			StringBuilder sb = new StringBuilder();
			CreatSkillBaseHeaderFile(repeatfieldType, sb);
			SaveToFile(sb, "CPP", @"\BaseSkill.h");
			CreatSkillBaseCppFile(repeatfieldType, sb);
			SaveToFile(sb, "CPP", @"\BaseSkill.cpp");
			
		}
		private void CreatSkillBaseHeaderFile(Dictionary<string, string> repeatfieldType, StringBuilder sb)
		{
			sb.Clear();
			sb.AppendLine("#pragma once");
			sb.AppendLine("#include \"DataFormat.h\"\n");
			sb.AppendLine("typedef uint32 uint;");
			sb.AppendLine("class BaseSkill");
			sb.AppendLine("{");
			sb.AppendLine("public:");
			sb.AppendLine("\tBaseSkill();");
			sb.AppendLine("\tvirtual ~BaseSkill();\n");
			sb.AppendLine("public:");
			foreach (var field in repeatfieldType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("\tconst string& Get" + field.Key + "()const;");
				}
				else
				{
					sb.AppendLine("\t" + field.Value + " Get" + field.Key + "()const;");
				}
			}
			sb.AppendLine("\tvirtual bool Release() = 0;");
			sb.AppendLine("\tvirtual bool Init(int level) = 0;\n");
			sb.AppendLine("protected:");
			foreach (var field in repeatfieldType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("\tvoid Set" + field.Key + "(const string& value);");
				}
				else
				{
					sb.AppendLine("\tvoid Set" + field.Key + "(" + field.Value + " value);");
				}
			}
			sb.AppendLine();
			sb.AppendLine("private:");
			foreach (var field in repeatfieldType)
			{
				sb.AppendLine("\t" + field.Value + " " + field.Key + ";");
			}
			sb.AppendLine("};");
		}
		private void CreatSkillBaseCppFile(Dictionary<string, string> repeatfieldType, StringBuilder sb)
		{
			sb.Clear();
			sb.AppendLine("#include \"BaseSkill.h\"\n");
			sb.AppendLine("BaseSkill::BaseSkill()");
			sb.AppendLine("{\n");
			sb.AppendLine("}\n");
			sb.AppendLine("BaseSkill::~BaseSkill()");
			sb.AppendLine("{\n");
			sb.AppendLine("}\n");

			//Get方法
			foreach (var field in repeatfieldType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("const std::string& BaseSkill::Get" + field.Key + "()const");
				}
				else
				{
					sb.AppendLine(field.Value + " BaseSkill::Get" + field.Key + "()const");
				}
				sb.AppendLine("{\n");
				sb.AppendLine("\treturn " + field.Key + ";");
				sb.AppendLine("}\n");
			}
			//Set方法
			foreach (var field in repeatfieldType)
			{
				if (field.Value.Equals("string"))
				{
					sb.AppendLine("void BaseSkill::Set" + field.Key + "(const string& value)");
				}
				else
				{
					sb.AppendLine("void BaseSkill::Set"  + field.Key + "("+field.Value +" value)");
				}
				sb.AppendLine("{\n");
				sb.AppendLine("\t " + field.Key + " = value;");
				sb.AppendLine("}\n");
			}
		}
		#endregion
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
				//生成各种代码
                foreach (var item in CodeCreater)
                {
					item.InitSheetName(validSheets);
                    item.ObjCodeCreat(validSheets);
					item.JsonReaderCodeCreat(validSheets);
					item.SkillCodeCreat(validSheets);
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
