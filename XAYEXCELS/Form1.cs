using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml;
using System.Net.Mail;
using System.Collections;
using System.Net;
using System.Threading;

namespace XAYEXCELS
{
    public partial class Form1 : Form
    {
        String arg;
        string log;
       public Form1(String[] args)
        {

            if (args.Length > 0)
            {
                //获取启动时的命令行参数  
                arg = args[0];
            }
            InitializeComponent();
        }
        IWorkbook workbook;
        public System.Data.DataTable ImportExcelFile(string filepath)
        {
            #region
            string[] targetfile = Directory.GetFiles(filepath);
            DataTable drtable = new System.Data.DataTable();
            for (int k = 0; k < targetfile.Length; k++)
            {
            try
            {
                string filePath = targetfile[k];
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //file.Position = 0;
                    if (filePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else /*if (filePath.EndsWith(".xlsx"))*/
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
                ISheet sheet = workbook.GetSheetAt(0);
                IRow headerRow = sheet.GetRow(0);
                int cellCount = 18;
                int rowCount = sheet.LastRowNum;
                if (k == 0)
                {             
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        bool samecol = drtable.Columns.Contains(headerRow.GetCell(i).ToString());
                        if (!samecol)
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                            drtable.Columns.Add(column);
                        }
                    }
                }

                for (int i = (sheet.FirstRowNum + 2); i <= rowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dataRow = drtable.NewRow();

                    if (row != null)
                    {
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                dataRow[j] = row.GetCell(j);
                            }

                        }
                    }

                    drtable.Rows.Add(dataRow);
                }
            }
 #endregion           
            return drtable;
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
        public static void ExportEasy(DataTable dtSource,string strFileName)
        {
            XSSFWorkbook workbook1 = new XSSFWorkbook();
            ISheet sheet = workbook1.CreateSheet();           
            IRow dataRow = sheet.CreateRow(0);
            ICellStyle style = workbook1.CreateCellStyle();
            ICellStyle styleh = workbook1.CreateCellStyle();
            IFont font = workbook1.CreateFont();
            dataRow.Height = 28 * 20;     
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;             
            font.FontName = "宋体";           
            font.FontHeightInPoints = 10;
            style.SetFont(font);
            styleh.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleh.VerticalAlignment = VerticalAlignment.Center;
            styleh.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightTurquoise.Index;
            styleh.FillPattern = FillPattern.SolidForeground;
            styleh.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.LightTurquoise.Index;
            //font.IsBold =true;
            styleh.SetFont(font);            
            foreach (DataColumn column in dtSource.Columns)
            {
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                dataRow.Cells[column.Ordinal].CellStyle = styleh;
            }


            //填充内容
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dtSource.Rows[i][j].ToString());
                    dataRow.Cells[j].CellStyle = style;
                   
                }
              
            }

           
            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook1.Write(fs);
                }
            }
            
        }
        public void MultiSendEmail(object data)
        {
            string emailstr = data as string;
            string[] emaildata = emailstr.Split('☆');
            try
            {            
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            mail.To = emaildata[0];
            mail.Attachments.Add(emaildata[1], Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mail.Subject = emaildata[2];
            mail.Body = emaildata[3];
            //mail.Display(true);
            mail.Send();
            mail = null;
            app = null;
                log =log+ DateTime.Now.ToLongTimeString()+"  "+ emaildata[4] + "的邮件发送成功\r\n";


            }
            catch (System.Exception ex)
            {
                notifyIcon1.ShowBalloonTip(4000, "提示", emaildata[4]+ "的邮件发送失败", ToolTipIcon.None);
                log =log+DateTime.Now.ToLongTimeString() +"  "+ emaildata[4] + "的邮件发送失败\r\n";
            }

        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.notifyIcon1.Visible = true;
                this.ShowInTaskbar = false;
            }
        }            
     
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            this.Close();
        }     
        private void closetsm_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void runtsm_Click(object sender, EventArgs e)
        {
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            DataTable option = ReadFromXml(sysstr+ "\\Option.xml");
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();               
             DataTable dt= ImportExcelFile(option.Rows[0].ItemArray[0].ToString());            
            dt.DefaultView.Sort = "业务员,产品名称,数量";
            dt = dt.DefaultView.ToTable();
            string strFileName = option.Rows[1].ItemArray[0].ToString() +"\\"+ DateTime.Now.ToString("yyMMdd") + ".xlsx";
            string autoemail = option.Rows[2].ItemArray[0].ToString();
            ExportEasy(dt, strFileName);
            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            double seconds = timeSpan.TotalSeconds;
            log = log + DateTime.Now.ToLongTimeString() + "  共合并" + dt.Rows.Count + "行" + "，耗时" + seconds + "秒\r\n";
            DataTable dailidt = ReadFromXml(sysstr + "\\DataTable.xml");           
            for (int k = 0; k < dailidt.Rows.Count; k++)
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(MultiSendEmail));
                string daili = dailidt.Rows[k][0].ToString();
                string email= dailidt.Rows[k][1].ToString();
                string emailSubject = dailidt.Rows[k][2].ToString();
                string emailBody = dailidt.Rows[k][3].ToString();
                DataRow[] drArr = dt.Select("业务员 LIKE '"+ daili + "%'");
                DataTable dtNew = dt.Clone();
                if (drArr.Length == 0)               
                    continue;               
                for (int i = 0; i < drArr.Length; i++)
                {
                    dtNew.ImportRow(drArr[i]);

                }
                if (autoemail == "True")
                {
                    strFileName = option.Rows[1].ItemArray[0].ToString() + "\\" + daili + DateTime.Now.ToString("yyMMdd") + ".xlsx";
                    string data = email + "☆" + strFileName + "☆" + emailSubject + "☆" + emailBody + "☆" + daili;
                    ExportEasy(dtNew, strFileName);
                    t1.IsBackground = true;
                    t1.Start(data);
                }
                
            }            
            textBox1.Text = log;
            //notifyIcon1.ShowBalloonTip(3000,"提示", textBox1.Text, ToolTipIcon.None);


        }

        private void option_Click(object sender, EventArgs e)
        {
            new Form2().Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (arg != null)
            {
                this.Visible = false;
                this.ShowInTaskbar = false;
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = log;
        }
    }
   
  }
