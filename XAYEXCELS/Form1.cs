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

namespace XAYEXCELS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        IWorkbook workbook;
        public System.Data.DataTable ImportExcelFile()
        {
            #region
            string[] targetfile = Directory.GetFiles("F:\\代理");
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
        public static void ExportEasy(DataTable dtSource)
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
            string strFileName = "F:\\代理\\生成\\" + DateTime.Now.ToString("yyMMdd")+".xlsx";
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

       
        //protected override void OnClosing(CancelEventArgs e)
        //{
        //    this.ShowInTaskbar = false;
        //    this.WindowState = FormWindowState.Minimized;
        //    e.Cancel = true;
        //}

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

      

        private void inputtsm_Click(object sender, EventArgs e)
        {

        }

        private void outputtsm_Click(object sender, EventArgs e)
        {

        }

        private void closetsm_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void runtsm_Click(object sender, EventArgs e)
        {
           
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();
            string[] a = System.IO.Directory.GetFiles("F:\\代理");          
             DataTable dt= ImportExcelFile();            
            dt.DefaultView.Sort = "业务员,产品名称,数量";
            dt = dt.DefaultView.ToTable();
            ExportEasy(dt);
            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            double seconds = timeSpan.TotalSeconds;
            textBox1.Text = "共合并" + dt.Rows.Count + "行"+"，耗时"+seconds+"秒";


        }

    }
   
  }
