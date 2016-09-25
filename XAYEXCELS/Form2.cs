using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace XAYEXCELS
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
          
            
        }
        public static bool WriteToXml(DataTable dt, string address)
        {
            try
            {
                //如果文件DataTable.xml存在则直接删除
                if (File.Exists(address))
                {
                    File.Delete(address);
                }
                XmlTextWriter writer =
                 new XmlTextWriter(address, Encoding.GetEncoding("GBK"));
                writer.Formatting = Formatting.Indented;
                //XML文档创建开始
                writer.WriteStartDocument();
                writer.WriteComment("DataTable: " + dt.TableName);
                writer.WriteStartElement("DataTable"); //DataTable开始
                writer.WriteAttributeString("TableName", dt.TableName);
                writer.WriteAttributeString("CountOfRows", dt.Rows.Count.ToString());
                writer.WriteAttributeString("CountOfColumns", dt.Columns.Count.ToString());
                writer.WriteStartElement("ClomunName", ""); //ColumnName开始
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    writer.WriteAttributeString(
                     "Column" + i.ToString(), dt.Columns[i].ColumnName);
                }
                writer.WriteEndElement(); //ColumnName结束
                                          //按行各行
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    writer.WriteStartElement("Row" + j.ToString(), "");
                    //打印各列
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        writer.WriteAttributeString(
                         "Column" + k.ToString(), dt.Rows[j][k].ToString());
                    }
                    writer.WriteEndElement();
                }
                writer.WriteEndElement(); //DataTable结束
                writer.WriteEndDocument();
                writer.Close();
                //XML文档创建结束
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }
        public static DataTable ReadFromXml(string address)
        {
            DataTable dt = new DataTable();
            try
            {
                if (!File.Exists(address))
                {
                    throw new Exception("文件不存在!");
                }
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(address);
                XmlNode root = xmlDoc.SelectSingleNode("DataTable");
                //读取表名
                dt.TableName = ((XmlElement)root).GetAttribute("TableName");
                //Console.WriteLine("读取表名： {0}", dt.TableName);
                //读取行数
                int CountOfRows = 0;
                if (!int.TryParse(((XmlElement)root).
                 GetAttribute("CountOfRows").ToString(), out CountOfRows))
                {
                    throw new Exception("行数转换失败");
                }
                //读取列数
                int CountOfColumns = 0;
                if (!int.TryParse(((XmlElement)root).
                 GetAttribute("CountOfColumns").ToString(), out CountOfColumns))
                {
                    throw new Exception("列数转换失败");
                }
                //从第一行中读取记录的列名
                foreach (XmlAttribute xa in root.ChildNodes[0].Attributes)
                {
                    dt.Columns.Add(xa.Value);
                    //Console.WriteLine("建立列： {0}", xa.Value);
                }
                //从后面的行中读取行信息
                for (int i = 1; i < root.ChildNodes.Count; i++)
                {
                    string[] array = new string[root.ChildNodes[0].Attributes.Count];
                    for (int j = 0; j < array.Length; j++)
                    {
                        array[j] = root.ChildNodes[i].Attributes[j].Value.ToString();
                    }
                    dt.Rows.Add(array);
                    //Console.WriteLine("行插入成功");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return new DataTable();
            }
            return dt;
        }
       

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            WriteToXml(dt, "DataTable.xml");             
            DataTable dt2 = new DataTable("option");
            WriteToXml(dt2, "DataTableSET.xml");           
            dt2.Columns.Add("value");           
            dt2.Rows.Add(TBY.Text);
            dt2.Rows.Add(TYO.Text);
            WriteToXml(dt2, "DataTableSET.xml");
           


        }

        private void Form2_Load(object sender, EventArgs e)
        {
            DataTable dt = ReadFromXml("DataTableSET.xml");
            TBY.Text = dt.Rows[0].ItemArray[0].ToString();
            TYO.Text = dt.Rows[1].ItemArray[0].ToString();
            DataTable dt2 = ReadFromXml("DataTable.xml");
            dataGridView1.DataSource = dt2;
           
        }    
     
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DataTable dt = dataGridView1.DataSource as DataTable;          
            WriteToXml(dt, "DataTable.xml");           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            TBY.Text=folderBrowserDialog1.SelectedPath ;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            TYO.Text = folderBrowserDialog1.SelectedPath;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked) //设置开机自启动  
            {
                string path = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                rk2.SetValue("JcShutdown", path + " -s");
                rk2.Close();
                rk.Close();
            }
            else //取消开机自启动  
            {              
                string path = Application.ExecutablePath;
                RegistryKey rk = Registry.LocalMachine;
                RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                rk2.DeleteValue("JcShutdown", false);
                rk2.Close();
                rk.Close();
            }
        }
    }
}
