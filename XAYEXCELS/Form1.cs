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
using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using NPOI.HPSF;
using System.Globalization;
using System.Net.Sockets;
using System.Text.RegularExpressions;

namespace XAYEXCELS
{
    public partial class Form1 : Form
    {
        String arg;
        string log;
        string[] restartemail=new string[30];
        string[] emailarg = new string[1000];
        int time1;
        int emailtime;
       public Form1(String[] args)
        {
           
            if (args.Length > 0)
            {
                //获取启动时的命令行参数  
                arg = args[0];
            }
            InitializeComponent();
            Thread pipethread = new Thread(new ThreadStart(receiveStream));
            pipethread.IsBackground = true;
            pipethread.Start();
            


        }

        private void receiveStream()
        {
            try
            {
                while (true)
                {

                    TcpListener listener = new TcpListener(new IPEndPoint(IPAddress.Any, 1298));
                    listener.Start();
                    TcpClient remoteClient = listener.AcceptTcpClient();

                    NetworkStream streamToClient = remoteClient.GetStream();
                    lock (streamToClient)
                    {
                        byte[] bytes = Read2Buffer(streamToClient, 1024 * 1000);
                        BinaryFormatter bf = new BinaryFormatter();
                        streamToClient.Read(bytes, 0, bytes.Length);
                        streamToClient.Close();
                        MemoryStream ms = new MemoryStream(bytes);
                        DataTable dt = bf.Deserialize(ms) as DataTable;
                        for (int i=0;i<dt.Rows.Count;i++)
                        {
                            string str = dt.Rows[i].ItemArray[13].ToString();
                            string str2;
                            str2 = str.Substring(0, str.Split(' ')[0].Length + str.Split(' ')[1].Length + str.Split(' ')[2].Length + 3);
                            str = str.Replace(str2, "");
                            str2 = str2.Replace(" ", "");
                            dt.Rows[i]["地址"] = str2 + str;
                            string product = dt.Rows[i]["产品名称"].ToString();
                            string Logistical = dt.Rows[i]["物流公司"].ToString();
                            dt.Rows[i]["产品名称"] = product.Split('（')[0];
                            dt.Rows[i]["物流公司"] = Logistical.Split('（')[0];
                        }
                        ExcelExport(dt);
                        listener.Stop();
                        Thread.Sleep(1000);

                    }
                   
                }
               
            }
            catch(System.Exception ex)
            {
                notifyIcon1.ShowBalloonTip(2000, "提示", ex.Message, ToolTipIcon.Error);
            }
        }

        IWorkbook workbook;

        public static byte[] Read2Buffer(Stream stream, int BufferLen)
        {
            // 如果指定的无效长度的缓冲区，则指定一个默认的长度作为缓存大小
            if (BufferLen < 1)
            {
                BufferLen = 0x8000;
            }

            // 初始化一个缓存区
            byte[] buffer = new byte[BufferLen];
            int read = 0;
            int block;

            // 每次从流中读取缓存大小的数据，直到读取完所有的流为止
            while ((block = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                // 重新设定读取位置
                read += block;

                // 检查是否到达了缓存的边界，检查是否还有可以读取的信息
                if (read == buffer.Length)
                {
                    // 尝试读取一个字节
                    int nextByte = stream.ReadByte();

                    // 读取失败则说明读取完成可以返回结果
                    if (nextByte == -1)
                    {
                        return buffer;
                    }

                    // 调整数组大小准备继续读取
                    byte[] newBuf = new byte[buffer.Length * 2];
                    System.Array.Copy(buffer, newBuf, buffer.Length);
                    newBuf[read] = (byte)nextByte;

                    // buffer是一个引用（指针），这里意在重新设定buffer指针指向一个更大的内存
                    buffer = newBuf;
                    read++;
                }
            }

            // 如果缓存太大则使用ret来收缩前面while读取的buffer，然后直接返回
            byte[] ret = new byte[read];
            System.Array.Copy(buffer, ret, read);
            return ret;
        }
        public System.Data.DataTable ImportExcelFile(string filepath,string mode)
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
                IRow headerRow;
                int cellCount;
                int rowCount;
                if (mode == "0")
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = 18;
                    rowCount = sheet.LastRowNum;
                 
                }
                else
                {
                    headerRow = sheet.GetRow(0);
                    cellCount = 31;
                    rowCount = sheet.LastRowNum;
                   
                }
                if (k == 0)
                {             
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        bool samecol = drtable.Columns.Contains(headerRow.GetCell(i).ToString());
                        if (!samecol)
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                            if (column.ColumnName==null) { column.ColumnName = ""; }
                            drtable.Columns.Add(column);
                            //if ((i == 1 || i == 12) && mode == "1")
                            //{
                            //    column.DataType = System.Type.GetType("System.DateTime");
                            //}
                        }
                    }
                    
                }
                //try
                //{

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
                                    if (j == 0&& mode=="0" && row.GetCell(j).ToString() != "")
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue.ToShortDateString();
                                    }
                                    else if ((j == 1||j==22||j==26 )&& mode == "1" && row.GetCell(j).ToString() != "" && row.GetCell(j).ToString() != "购买日期"&& row.GetCell(j).CellType != CellType.String)
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
                                    else if (row.GetCell(j).CellType == CellType.Formula && (j == 5 || j == 16))
                                    {
                                        dataRow[j] = row.GetCell(j).NumericCellValue;
                                    }                                    
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j);
                                    }                       

                                }

                            }
                        }
                   
                    drtable.Rows.Add(dataRow);
                }
                //}
                //catch (Exception ex)
                //{
                //    notifyIcon1.ShowBalloonTip(2000, "提示", "运行失败，excel文件单元格格式有误", ToolTipIcon.Error);
                //    log = log  +DateTime.Now.ToLongTimeString() + "  运行失败，excel文件单元格格式有误\r\n";
                //    break;
                //}

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
            HSSFWorkbook workbook1 = new HSSFWorkbook();
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
                        else if (exceltype == "3" && (j == 2 || j == 5 || j == 6) && dtcellvalue != "")
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
            sheet.CreateFreezePane(3, 1, 3, 1);
                CellRangeAddress range = CellRangeAddress.ValueOf("A1:Q1");
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
        public void ExportEasySH(DataTable dtSource, string strFileName, string exceltype)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            #region 右击文件 属性信息
            {
                NPOI.HPSF.DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "CANDA"; 
                si.ApplicationName = "NPOI";
                si.LastAuthor = "CANDA"; 
                si.Comments = "CANDA";
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            
            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            //int[] arrColWidth = new int[dtSource.Columns.Count];
            //foreach (DataColumn item in dtSource.Columns)
            //{
            //    arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            //}
            //for (int i = 0; i < dtSource.Rows.Count; i++)
            //{
            //    for (int j = 0; j < dtSource.Columns.Count; j++)
            //    {
            //        int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
            //        if (intTemp > arrColWidth[j])
            //        {
            //            arrColWidth[j] = intTemp;
            //        }
            //    }
            //}
         
            int rowIndex = 0;

            foreach (DataRow row in dtSource.Rows)
            {
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }
                    #region 列头及样式
                    {
                        IRow headerRow = sheet.CreateRow(0);                        
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        headStyle.VerticalAlignment = VerticalAlignment.Center;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;
                            //sheet.SetColumnWidth(column.Ordinal,(arrColWidth[column.Ordinal]) * 256);
                        }
                    }
                    #endregion

                    rowIndex = 1;
                }
             

                #region 填充内容
                IRow dataRow = sheet.CreateRow(rowIndex);
                ICellStyle style = workbook.CreateCellStyle();
                style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
               
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);
                    string drValue = row[column].ToString();
                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型
                            newCell.SetCellValue(drValue);
                            newCell.CellStyle = style;
                            break;
                        case "System.DateTime"://日期类型
                            DateTime dateV;
                            IFormatProvider ifp = new CultureInfo("zh-CN", true);
                            DateTime.TryParseExact(drValue,"MM-dd",ifp,System.Globalization.DateTimeStyles.None,out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle;//格式化显示
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }
                #endregion

                rowIndex++;
            }
            //sheet.ForceFormulaRecalculation = true;
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 0));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 1, 1));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 2, 2));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 3, 3));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 4, 4));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 5, 5));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 6, 6));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 7, 7));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 8, 11));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 12, 12));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 13, 13));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 14, 14));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 16, 16));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 15, 15));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 17, 19));
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 20, 20));
            //sheet.SetDefaultColumnStyle(1,);
            sheet.SetColumnWidth(4, 13 * 256);
            sheet.SetColumnWidth(5, 20 * 256);
            sheet.SetColumnWidth(7, 20 * 256);
            sheet.SetColumnWidth(13, 48 * 256);
            sheet.SetColumnWidth(15, 15 * 256);
            sheet.SetColumnWidth(15, 19 * 256);
            sheet.CreateFreezePane(11, 0, 11, 0);
            CellRangeAddress range = CellRangeAddress.ValueOf("A2:U2");
            sheet.SetAutoFilter(range);

            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }

        }

        public void MultiSendEmail(object emailarr)
        {
            string[] data = emailarr as string[];
            for(int i = 0; i < data.Length; i++)
            {
            string emailstr = data[i];
                if (emailstr == null)
                    continue;
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
                notifyIcon1.ShowBalloonTip(2000, "提示", emaildata[4]+ "的邮件发送失败"+","+ex.Message, ToolTipIcon.None);
                log =log+DateTime.Now.ToLongTimeString() +"  "+ emaildata[4] + "的邮件发送失败\r\n";
                restartemail[time1] = data[i];
                time1++;
            }
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
            notifyIcon1.Visible = false;
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
            string excelmode= (option.Rows[4].ItemArray[0].ToString()=="True")?"0":"1";
            string daydir = filepath + "\\" + DateTime.Now.Date.ToString("yyyy.MM") + "\\" + DateTime.Now.Date.ToString("dd") + "日";
            string targetdir = daydir + "\\" + "源文件";
            if (!Directory.Exists(targetdir))
                Directory.CreateDirectory(targetdir);
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
            DataTable dt = ImportExcelFile(targetdir,excelmode);
            if (dt.Rows.Count == 0)
            {
                log = log + DateTime.Now.ToLongTimeString() + "  源文件为空\r\n";
                textBox1.Text = log;
                notifyIcon1.ShowBalloonTip(3000, "提示", "请添加源文件", ToolTipIcon.None);
                return;
            }
            if (excelmode == "0")
            {
                ExcelExport(dt);
            }
            else if (excelmode == "1")
            {
                ExcelExportSH(dt);
            }
            
            textBox1.Text = log;

        }

        private void ExcelExport(DataTable dt)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            DataTable option = ReadFromXml(sysstr + "\\Option.xml");
            string filepathInput = option.Rows[0].ItemArray[0].ToString();
            string filepath = option.Rows[1].ItemArray[0].ToString();
            string daydir = filepath + "\\" + DateTime.Now.Date.ToString("yyyy.MM") + "\\" + DateTime.Now.Date.ToString("dd") + "日";           
            string dailidir = daydir + "\\" + "单号表";
            string totaldir = daydir + "\\" + "汇总表";
            emailtime = 0;            
            if (!Directory.Exists(dailidir))
                Directory.CreateDirectory(dailidir);
            if (!Directory.Exists(totaldir))
                Directory.CreateDirectory(totaldir);           
            dt.DefaultView.Sort = "业务员 DESC";
            dt = dt.DefaultView.ToTable();
            string strFileName = totaldir + "\\汇总表" + DateTime.Now.ToString("yyMMdd") + ".xls";
            string autoemail = option.Rows[2].ItemArray[0].ToString();
            
            ExportEasy(dt, strFileName, "1");
            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            double seconds = timeSpan.TotalSeconds;
            log = log + DateTime.Now.ToLongTimeString() + "  共合并" + dt.Rows.Count + "行" + "，耗时" + seconds + "秒\r\n";
            DataTable dailidt = ReadFromXml(sysstr + "\\DataTable.xml");
            string[] emailarr = new string[dailidt.Rows.Count];
            for (int k = 0; k < dailidt.Rows.Count; k++)
            {
                string daili = dailidt.Rows[k][0].ToString();
                string email = dailidt.Rows[k][1].ToString();
                string email2 = dailidt.Rows[k][2].ToString();
                string emailSubject = dailidt.Rows[k][3].ToString();
                string emailBody = dailidt.Rows[k][4].ToString();
                string daili2 = dailidt.Rows[k][5].ToString();
                DataRow[] drArr;
                if (daili2 == "")
                {
                    drArr = dt.Select("业务员 LIKE '" + daili.Substring(0, 1) + "%'");
                }
                else
                {
                    drArr = dt.Select("业务员 LIKE '" + daili2.Substring(0, 1) + "%' and 业务员 Like '%" + daili + "%' ");
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
                //dtNew.Columns.Remove("注释");
                strFileName = dailidir + "\\" + emailSubject + "(" + DateTime.Now.ToString("M.dd") + ").xls";
                ExportEasy(dtNew, strFileName, "2");
                string data = email + "☆" + strFileName + "☆" + emailSubject + "☆" + emailBody + "☆" + daili + "☆" + email2;
                emailarg[emailtime] = data;
                emailarr[k] = data;
                emailtime++;
            }
            if (autoemail == "True")
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(MultiSendEmail));
                t1.IsBackground = true;
                t1.Start(emailarr);
                //MultiSendEmail(data);
            }
            //notifyIcon1.ShowBalloonTip(3000,"提示", textBox1.Text, ToolTipIcon.None);
        }
        private void ExcelExportSH(DataTable dt)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            DataTable option = ReadFromXml(sysstr + "\\Option.xml");
            string filepathInput = option.Rows[0].ItemArray[0].ToString();
            string filepath = option.Rows[1].ItemArray[0].ToString();
            string daydir = filepath + "\\" + DateTime.Now.Date.ToString("yyyy.MM") + "\\" + DateTime.Now.Date.ToString("dd") + "日";
            string dailidir = daydir + "\\" + "单号表";
            string totaldir = daydir + "\\" + "汇总表";
            emailtime = 0;
            if (!Directory.Exists(dailidir))
                Directory.CreateDirectory(dailidir);
            if (!Directory.Exists(totaldir))
                Directory.CreateDirectory(totaldir);
            //dt.Columns[8].ColumnName = "产品型号";
            //dt.Columns[9].ColumnName = "故障类型";
            //dt.Columns[10].ColumnName = "其它";
            //dt.Columns[11].ColumnName = "处理方式";
            dt.Columns[26].ColumnName = "购买日期";
            dt.Columns.Remove("处 理 结 果");
            dt.Columns.Remove("Column3");
            dt.Columns.Remove("Column4");
            dt.Columns.Remove("Column5");
            dt.Columns.Remove("Column6");
            dt.Columns.Remove("Column7");
            dt.Columns.Remove("Column8");
            dt.Columns.Remove("Column9");
            dt.Columns.Remove("Column10");
            dt.Columns.Remove("Column11");
           
            //dt.Columns[17].ColumnName = "快寄类型";
            //dt.Columns[18].ColumnName = "补贴";
            //dt.Columns[19].ColumnName = "单号";
            string strFileName = totaldir + "\\汇总表" + DateTime.Now.ToString("yyMMdd") + ".xls";
            string autoemail = option.Rows[2].ItemArray[0].ToString();
            ExportEasySH(dt, strFileName, "1");
            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            double seconds = timeSpan.TotalSeconds;
            log = log + DateTime.Now.ToLongTimeString() + "  共合并" + dt.Rows.Count + "行" + "，耗时" + seconds + "秒\r\n";
            DataTable dailidt = ReadFromXml(sysstr + "\\DataTable.xml");
            for (int k = 0; k < dailidt.Rows.Count; k++)
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(MultiSendEmail));
                string daili = dailidt.Rows[k][0].ToString();
                string email = dailidt.Rows[k][1].ToString();
                string email2 = dailidt.Rows[k][2].ToString();
                string emailSubject = dailidt.Rows[k][3].ToString();
                string emailBody = dailidt.Rows[k][4].ToString();
                string daili2 = dailidt.Rows[k][5].ToString();
                DataRow[] drArr;
                if (daili2 == "")
                {
                    drArr = dt.Select("销售员 LIKE '" + daili.Substring(0, 1) + "%' or 检测结果 = '产品型号'");
                }
                else
                {
                    drArr = dt.Select("(销售员 LIKE '" + daili2.Substring(0, 1) + "%' and 销售员 Like '%" + daili + "%') or 检测结果 = '产品型号' ");
                }

                DataTable dtNew = dt.Clone();
                if (drArr.Length == 1)
                    continue;
                
                for (int i = 0; i < drArr.Length; i++)
                {
                    dtNew.ImportRow(drArr[i]);
                }
                strFileName = dailidir + "\\" + emailSubject + "(" + DateTime.Now.ToString("M.dd") + ").xls";
                ExportEasySH(dtNew, strFileName, "0");
                string data = email + "☆" + strFileName + "☆" + emailSubject + "☆" + emailBody + "☆" + daili + "☆" + email2;
                emailarg[emailtime] = data;
                emailtime++;
                if (autoemail == "True")
                {
                    t1.IsBackground = true;
                    t1.Start(data);
                }

            }

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
            Thread t2 = new Thread(new ParameterizedThreadStart(MultiSendEmail));
            t2.IsBackground = true;
            t2.Start(restartemail);
          
        }

        private void sendbtn_Click(object sender, EventArgs e)
        {
            int time = emailtime;  
                Thread t = new Thread(new ParameterizedThreadStart(MultiSendEmail));
                t.IsBackground = true;
                t.Start(emailarg);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //注意判断关闭事件Reason来源于窗体按钮，否则用菜单退出时无法退出!
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;    //取消"关闭窗口"事件
                this.WindowState = FormWindowState.Minimized;    //使关闭时窗口向右下角缩小的效果
                notifyIcon1.Visible = true;
                this.Hide();
                return;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox1.Text = log;
        }
    }
   
  }
