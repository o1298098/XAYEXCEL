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
        string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
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


        private void button1_Click(object sender, EventArgs e)//保存
        {
            DataTable dt = dataGridView1.DataSource as DataTable;
            WriteToXml(dt, sysstr + "XAYXML\\DataTable.xml");
            dt = productGridView.DataSource as DataTable;
            WriteToXml(dt, sysstr + "XAYXML\\wldt.xml");
            DataTable dt2 = new DataTable("option");
            dt2.Columns.Add("value");
            dt2.Rows.Add(TBY.Text);
            dt2.Rows.Add(TYO.Text);
            dt2.Rows.Add(checkBox1.Checked);
            dt2.Rows.Add(checkBox2.Checked);
            dt2.Rows.Add(rbyewu.Checked);
            dt2.Rows.Add(rbshouhou.Checked);
            WriteToXml(dt2, sysstr + "XAYXML\\Option.xml");
            dt = comboBox1.DataSource as DataTable;
            WriteToXml(dt, sysstr + "XAYXML\\droplist.xml");
            dt = customerGridView.DataSource as DataTable;
            DataRowView selectitem = comboBox1.SelectedItem as DataRowView;
            WriteToXml(dt, sysstr + "XAYXML\\" + selectitem.Row.ItemArray[0] + ".xml");
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            DataTable dt = ReadFromXml(sysstr + "XAYXML\\Option.xml");
            TBY.Text = dt.Rows[0].ItemArray[0].ToString();
            TYO.Text = dt.Rows[1].ItemArray[0].ToString();
            if (dt.Rows[2].ItemArray[0].ToString() == "True")
            {
                checkBox1.Checked = true;
            }
            if (dt.Rows[3].ItemArray[0].ToString() == "True")
            {
                checkBox2.Checked = true;
            }
            if (dt.Rows[4].ItemArray[0].ToString() == "True")
            {
                rbyewu.Checked = true;
            }
            if (dt.Rows[5].ItemArray[0].ToString() == "True")
            {
                rbshouhou.Checked = true;
            }
            try
            {
                DataTable droplistdt = ReadFromXml(sysstr + "XAYXML\\droplist.xml");
                comboBox1.DataSource = droplistdt;
                comboBox1.DisplayMember = "chance";
                comboBox1.ValueMember = "chance";
                DataGridViewComboBoxColumn comUserName = new DataGridViewComboBoxColumn();
                comUserName.DataPropertyName = "代理方案";
                comUserName.HeaderText = "代理方案";
                comUserName.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                comUserName.DataSource = droplistdt;
                comUserName.DisplayMember = "chance";
                comUserName.ValueMember = "chance";
                dataGridView1.Columns.Add(comUserName);
                DataTable dtdali = ReadFromXml(sysstr + "XAYXML\\DataTable.xml");
                dataGridView1.DataSource = dtdali;
                DataTable dtwuliao = ReadFromXml(sysstr + "XAYXML\\wldt.xml");
                productGridView.DataSource = dtwuliao;
            }
            catch { }
           
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            DataTable dt = dataGridView1.DataSource as DataTable;
            WriteToXml(dt, sysstr + "XAYXML\\DataTable.xml");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            TBY.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            TYO.Text = folderBrowserDialog1.SelectedPath;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show("请在管理员模式下修改才能生效", "提示");

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void creatlink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DataTable dt = new DataTable("CUSTOMER");
            dt.Columns.Add("物料名称");
            dt.Columns.Add("物料编码");
            dt.Columns.Add("代理价");
            dt.Rows.Add("");
            string address = sysstr + "XAYXML\\" + comboBox1.Text + ".xml";
            if (!File.Exists(address))
            {
                WriteToXml(dt, address);
                dt = comboBox1.DataSource as DataTable;
                dt.Rows.Add(comboBox1.Text);
                comboBox1.DataSource = dt;
                customerGridView.DataSource = ReadFromXml(sysstr + "XAYXML\\" + comboBox1.Text + ".xml");
                dt = comboBox1.DataSource as DataTable;
                WriteToXml(dt, sysstr + "XAYXML\\droplist.xml");
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataRowView selectitem = comboBox1.SelectedItem as DataRowView;
            customerGridView.DataSource = ReadFromXml(sysstr + "XAYXML\\" + selectitem.Row.ItemArray[0] + ".xml");

        }

        private void deletelink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)//删除方案
        {
            DataRowView selectitem = comboBox1.SelectedItem as DataRowView;
            string address = sysstr + "XAYXML\\" + selectitem.Row.ItemArray[0] + ".xml";
            File.Delete(address);
            DataTable dt = comboBox1.DataSource as DataTable;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i].ItemArray[0] == selectitem.Row.ItemArray[0])
                {
                    dt.Rows.RemoveAt(i);
                }
            }
            comboBox1.DataSource = dt;
            WriteToXml(dt, sysstr + "XAYXML\\droplist.xml");
        }

        private void copylink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string address = sysstr + "XAYXML\\" + comboBox1.Text + ".xml";
            if (!File.Exists(address))
            {
                DataTable dt = customerGridView.DataSource as DataTable;
                WriteToXml(dt, sysstr + "XAYXML\\" + comboBox1.Text + ".xml");
                dt = comboBox1.DataSource as DataTable;
                dt.Rows.Add(comboBox1.Text);
                comboBox1.DataSource = dt;
                dt = comboBox1.DataSource as DataTable;
                WriteToXml(dt, sysstr + "XAYXML\\droplist.xml");
            }
        }
    }
}
