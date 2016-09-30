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
using NPOI.SS.Util;

namespace XAYEXCELS
{
    public partial class Form1 : Form
    {
        String arg;
        string log;
        string[] restartemail=new string[30];
        int time1;
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
            catch (Exception ex)
            {
                throw ex;
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
                try
                {

                    for (int i = (sheet.FirstRowNum + 1); i <= rowCount; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dataRow = drtable.NewRow();
                  if (row != null)
                        {
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    if (j == 0 && row.GetCell(j).ToString() != "")
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue.ToShortDateString();
                                    }
                                    else if (row.GetCell(j).CellType == CellType.Numeric)
                                    {
                                        dataRow[j] = row.GetCell(j).NumericCellValue;
                                    }
                                    else if (row.GetCell(j).CellType == CellType.String)
                                    {
                                        dataRow[j] = row.GetCell(j).StringCellValue.Trim();
                                    }
                                    else if (row.GetCell(j).CellType == CellType.Formula&&(j==5|| j == 16))
                                    {
                                        dataRow[j] = row.GetCell(j).NumericCellValue;
                                    }                                    

                                }

                            }
                        }
                   
                    drtable.Rows.Add(dataRow);
                }
                }
                catch (Exception ex)
                {
                    notifyIcon1.ShowBalloonTip(2000, "提示", "运行失败，excel文件单元格格式有误", ToolTipIcon.Error);
                    log = log  +DateTime.Now.ToLongTimeString() + "  运行失败，excel文件单元格格式有误\r\n";
                    break;
                }

            }

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
        public  void ExportEasy(DataTable dtSource,string strFileName,string exceltype)
        {
            XSSFWorkbook workbook1 = new XSSFWorkbook();
            ISheet sheet = workbook1.CreateSheet();           
            IRow dataRow = sheet.CreateRow(0);
            ICellStyle style = workbook1.CreateCellStyle();
            ICellStyle style2 = workbook1.CreateCellStyle();
            ICellStyle styleh = workbook1.CreateCellStyle();           
            IFont font = workbook1.CreateFont();
            dataRow.Height = 28 * 20;
            //sheet.DefaultColumnWidth = 10 * 256;
            sheet.SetColumnWidth(0, 10 * 256);
            sheet.SetColumnWidth(1, 11 * 256);
            sheet.SetColumnWidth(2, 9 * 256);
            if (exceltype == "2")
            {
                sheet.SetColumnWidth(11, 20 * 256);
            }
            sheet.SetColumnWidth(13, 13 * 256);
            sheet.SetColumnWidth(14, 15 * 256);
            sheet.SetColumnWidth(15, 13 * 256);
            sheet.SetColumnWidth(16, 15 * 256);           
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;            
            font.FontName = "宋体";           
            font.FontHeightInPoints = 10;
            style.SetFont(font);
            style2.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style2.VerticalAlignment = VerticalAlignment.Center;
            style2.SetFont(font);
            styleh.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleh.VerticalAlignment = VerticalAlignment.Center;
            styleh.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            styleh.FillPattern = FillPattern.SolidForeground;
            styleh.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            styleh.BorderLeft= NPOI.SS.UserModel.BorderStyle.Thin;
            styleh.BorderRight= NPOI.SS.UserModel.BorderStyle.Thin;
            //font.IsBold =true;
            styleh.SetFont(font);            
            foreach (DataColumn column in dtSource.Columns)
            {
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);               
                dataRow.Cells[column.Ordinal].CellStyle = styleh;
            }

            try
            {
                //填充内容
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    dataRow = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                    string dtcellvalue = dtSource.Rows[i][j].ToString();
                        if (exceltype == "1" && (j == 2 || j == 4 || j == 5 || j == 7 || j == 8) && dtcellvalue != "")
                        {

                            dataRow.CreateCell(j).SetCellValue(int.Parse(dtcellvalue));
                            dataRow.Cells[j].CellStyle = style2;
                           


                        }
                        else if (exceltype == "2" && (j == 2 || j == 5 || j == 6) && dtcellvalue != "")
                        {

                            dataRow.CreateCell(j).SetCellValue(int.Parse(dtcellvalue));
                            dataRow.Cells[j].CellStyle = style2;
                            
                        }
                        else
                        {
                            dataRow.CreateCell(j).SetCellValue(dtcellvalue);
                            dataRow.Cells[j].CellStyle = style;                           

                        }
                        //


                    }

                }
            }
            catch (Exception ex)
            {

                log = log + DateTime.Now.ToLongTimeString() + "  金额单元格格式有误\r\n";
                notifyIcon1.ShowBalloonTip(2000, "提示", "金额单元格格式有误", ToolTipIcon.Error);
                throw ex;
            }

            //sheet.ForceFormulaRecalculation = true;
            sheet.CreateFreezePane(3, 0, 4, 0);
                CellRangeAddress range = CellRangeAddress.ValueOf("A1:P1");
                sheet.SetAutoFilter(range);

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
                mail.To = emaildata[0]+";"+ emaildata[5];
                mail.Attachments.Add(emaildata[1], Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                mail.Subject = emaildata[2]+"("+DateTime.Now.Date.ToString("M.dd")+")";
                mail.Body = emaildata[3];
            //mail.Display(true);
                mail.Send();
                mail = null;
                app = null;
                log = log + DateTime.Now.ToLongTimeString() + "  " + emaildata[4] + "的邮件发送成功\r\n";             
               

            }
            catch (System.Exception ex)
            {
                notifyIcon1.ShowBalloonTip(2000, "提示", emaildata[4]+ "的邮件发送失败", ToolTipIcon.None);
                log =log+DateTime.Now.ToLongTimeString() +"  "+ emaildata[4] + "的邮件发送失败\r\n";
                restartemail[time1] = data as string;
                time1++;
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
            DataTable option = ReadFromXml(sysstr + "\\Option.xml");
            string filepathInput = option.Rows[0].ItemArray[0].ToString();
            string filepath = option.Rows[1].ItemArray[0].ToString();            
            string daydir = filepath + "\\" + DateTime.Now.Year.ToString() + "\\" + DateTime.Now.Month.ToString("") + "月" + "\\" + DateTime.Now.Date.ToString("M.dd");
            string targetdir = daydir + "\\" + "源文件";
            string dailidir = daydir + "\\" + "单号表";
            string totaldir = daydir + "\\" + "汇总表";
            if (!Directory.Exists(targetdir))
                Directory.CreateDirectory(targetdir);
            if (!Directory.Exists(dailidir))
                Directory.CreateDirectory(dailidir);
            if (!Directory.Exists(totaldir))
                Directory.CreateDirectory(totaldir);
            List<string> files = new List<string>(Directory.GetFiles(filepathInput));
            files.ForEach(c =>
            {
                string destFile = Path.Combine(new string[] { targetdir, Path.GetFileName(c) });              
                if (File.Exists(destFile))
                {
                    File.Delete(destFile);
                }
                File.Move(c, destFile);
            });
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();               
            DataTable dt= ImportExcelFile(targetdir);
            if (dt.Rows.Count == 0)
            {
                log = log + DateTime.Now.ToLongTimeString() + "  源文件为空\r\n";
                textBox1.Text = log;
                notifyIcon1.ShowBalloonTip(3000, "提示","请添加源文件", ToolTipIcon.None);
                return;
            }

            dt.DefaultView.Sort = "业务员 DESC";
            dt = dt.DefaultView.ToTable();
            string strFileName = totaldir + "\\汇总表"+ DateTime.Now.ToString("yyMMdd") + ".xlsx";
            string autoemail = option.Rows[2].ItemArray[0].ToString();
            ExportEasy(dt, strFileName,"1");
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
                string email2 = dailidt.Rows[k][2].ToString();
                string emailSubject = dailidt.Rows[k][3].ToString();
                string emailBody = dailidt.Rows[k][4].ToString();
                string daili2= dailidt.Rows[k][5].ToString();
                DataRow[] drArr;
                if (daili2 == "")
                {
                     drArr = dt.Select("业务员 LIKE '" + daili.Substring(0, 1) + "%'");
                }
                else
                {
                    drArr = dt.Select("业务员 LIKE '" + daili2.Substring(0, 1) + "%' and 业务员 Like '%" + daili+ "%' ");
                }
               
                DataTable dtNew = dt.Clone();
                if (drArr.Length == 0)               
                    continue;               
                for (int i = 0; i < drArr.Length; i++)
                {
                    dtNew.ImportRow(drArr[i]);
                }
                dtNew.Columns.Remove("结算费用");
                dtNew.Columns.Remove("拿货单价");
                dtNew.Columns.Remove("注释");
                strFileName = dailidir + "\\" + daili +"单号(" +DateTime.Now.ToString("M.dd") + ").xlsx";
                ExportEasy(dtNew, strFileName,"2");
                if (autoemail == "True")
                {                  
                    string data = email + "☆" + strFileName + "☆" + emailSubject + "☆" + emailBody + "☆" + daili + "☆" +email2;
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

        private void restart_Click(object sender, EventArgs e)
        {
            
            int time2 = time1;
            time1 = 0;
            for (int i = 0; i < time2; i++)
            {
                Thread t2 = new Thread(new ParameterizedThreadStart(MultiSendEmail));
                t2.IsBackground = true;
                t2.Start(restartemail[i]);
            }
          
        }
    }
   
  }
