using System.Windows;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Windows.Media;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Text;
using Newtonsoft.Json;
using System;
using ExcelToJson.Properties;
using System.Diagnostics;

namespace ExcelToJson
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        List<FileInfo> files = new List<FileInfo>();
        public MainWindow()
        {
            InitializeComponent();
            SrcPath.Text = Directory.GetCurrentDirectory();

		}
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SrcPath.Text = fbd.SelectedPath;
            }
        }

        private void SrcPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            LogList.Items.Clear();
            files.Clear();
            //判断路径是否正确
            bool isRightPath = Directory.Exists(SrcPath.Text);
            if (isRightPath)
            {
                DirectoryInfo mydir = new DirectoryInfo(SrcPath.Text);

                //查找所有Excel文件
                files.AddRange(mydir.GetFiles("*.xls"));
                files.RemoveAll(s => s.ToString().Contains("~$"));
                foreach (var item in files)
                {
                    System.Windows.Controls.ListViewItem list = new System.Windows.Controls.ListViewItem();
                    list.Content = item;
                    LogList.Items.Add(list);
                }
                if (LogText != null)
                    LogText.Text = " ";
            }
            else
            {
                LogText.Text = "文件路径不正确!";
                LogText.Foreground = Brushes.Red;
                return;
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TargetPath.Text = fbd.SelectedPath;
            }

        }

        private void TargetPath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!Directory.Exists(TargetPath.Text))
            {
                LogText.Text = "目标路径不正确!";
                LogText.Foreground = Brushes.Red;
            }
            else
            {
                LogText.Text = "";
            }
            
        }
        private bool CheckPath()
        {
            //检查目标路径
            if (!Directory.Exists(SrcPath.Text))
            {
                LogText.Text = "文件路径不正确!";
                LogText.Foreground = Brushes.Red;
                return false;
            }
            //检查目标路径
            if (!Directory.Exists(TargetPath.Text))
            {
                LogText.Text = "目标路径不正确!";
                LogText.Foreground = Brushes.Red;
                return false;
            }
            return true;
        }
        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
			LogText.Text = "执行中......";
			LogText.Foreground = Brushes.Green;
			try
            {
                if (!CheckPath())
                    return;
                string strConn;
                DataSet ds = new DataSet();
				Stopwatch sw1 = new Stopwatch();
				Stopwatch sw2 = new Stopwatch();
				foreach (var file in files)
                {
                    List<string> al = new List<string>();
                    #region Get Sheets Name
                    strConn = "Provider=Microsoft.Ace.OleDb.12.0;Persist Security Info=False; data source="
                        + @file.FullName + ";Extended Properties='Excel 8.0; HDR=yes; IMEX=1'";
					sw1.Start();
                    using (OleDbConnection conn = new OleDbConnection(strConn))
                    {
                        conn.Open();
                        DataTable sheetNames = conn.GetOleDbSchemaTable
                        (OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        foreach (DataRow dr in sheetNames.Rows)
                        {
                            al.Add(dr[2].ToString());
                        }
                    }
					sw1.Stop();
					#endregion
					#region 将Sheet中的数据赋值到DataSet
					sw2.Start();
                    if ((bool)IsAllWorkBookCheckBox.IsChecked)
                    {
                        foreach (var item in al)
                        {
                            OleDbDataAdapter oada = new OleDbDataAdapter("select * from [" + item + "]", strConn);
                            //处理SheetName中的字符
                            string tableName = item.ToString().Replace("'", "");
                            tableName = tableName.Substring(0, tableName.Length - 1);
                            oada.Fill(ds, tableName);
                        }
                    }
                    else
                    {
                        OleDbDataAdapter oada = new OleDbDataAdapter("select * from [" + al[0] + "]", strConn);
                        //处理SheetName中的字符
                        string tableName = al[0].ToString().Replace("'", "");
                        tableName = tableName.Substring(0, tableName.Length - 1);
                        oada.Fill(ds, tableName);
                    }
					sw2.Stop();
                    #endregion
                }
				TimeSpan ts1 = sw1.Elapsed;
				Console.WriteLine("读取Excel总共花费{0}ms.", ts1.TotalMilliseconds);
				TimeSpan ts2 = sw2.Elapsed;
				Console.WriteLine("转入DataSet总共花费{0}ms.", ts2.TotalMilliseconds);
				WriteJson(ds);
				WriteCode(ds);

				LogText.Text = "转换成功！！";
                LogText.Foreground = Brushes.Green;
                System.Diagnostics.Process.Start(TargetPath.Text);
            }
            catch (Exception ex)
            {
                LogText.Text = "转换失败："+ex.ToString();
                LogText.Foreground = Brushes.Red;
            }
        }
        //写为Json
        private void WriteJson(DataSet ds)
        {
            Encoding utf8 = new UTF8Encoding(false);
            bool isArray = TransformTypeCombo.SelectedIndex == 0 ? true : false;
            var jsonExprter = new JsonExporter(ds, false, isArray, "yyyy/MM/dd", TargetPath.Text, utf8);
        }
        //写为代码
        private void WriteCode(DataSet ds)
        {
            Encoding utf8 = new UTF8Encoding(false);
            string path = TargetPath.Text + @"\Code";
            CodeCreaterManager codeCreatmanager = new CodeCreaterManager(path,utf8);
            if((bool)CreatCPPCode.IsChecked)
            {
				Stopwatch sw = new Stopwatch();
				sw.Start();
				codeCreatmanager.AddCreatCodeType(CreateType.CPP);
				sw.Stop();
				TimeSpan ts2 = sw.Elapsed;
				Console.WriteLine("写入C++总共花费{0}ms.", ts2.TotalMilliseconds);
			}
            if ((bool)CreatCSharpCode.IsChecked)
            {
				Stopwatch sw = new Stopwatch();
				sw.Start();
				codeCreatmanager.AddCreatCodeType(CreateType.CSharp);
				sw.Stop();
				TimeSpan ts2 = sw.Elapsed;
				Console.WriteLine("写入C#总共花费{0}ms.", ts2.TotalMilliseconds);
			}
            codeCreatmanager.CodeCreat(ds);
        }
    }
}
