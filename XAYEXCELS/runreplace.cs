using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace XAYEXCELS
{
    class runreplace
    {
        public string productreplace( string id,string name)
        {
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            System.Data.DataTable productdt = Form2.ReadFromXml(sysstr + "XAYXML\\wldt.xml");
            string product=name;
            for (int i = 0; i < productdt.Rows.Count; i++)
            {
                if (id == productdt.Rows[i].ItemArray[1].ToString())
                {
                    product = productdt.Rows[i].ItemArray[0].ToString();
                }
            }
            return product;
        }
        public System.Data.DataTable replacedt(System.Data.DataTable dt)
        {
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            string plan;
            System.Data.DataTable productdt = Form2.ReadFromXml(sysstr + "XAYXML\\DataTable.xml");
           
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sales = dt.Rows[i]["业务员"].ToString();
                DataRow[] dr = productdt.Select("代理名字 like '" + sales.Substring(0, 1) + "%'");
                dt.Rows[i]["代理单价"] = dt.Rows[i]["拿货单价"].ToString();
                if (dr.Length>0)
                {
                    plan = dr[0].ItemArray[6].ToString();
                    if(plan !="")
                    {
                        System.Data.DataTable plandt = Form2.ReadFromXml(sysstr + "XAYXML\\"+ plan + ".xml");
                    for (int j=0;j<plandt.Rows.Count;j++)
                    {
                            if (plandt.Rows[j].ItemArray[1].ToString()== dt.Rows[i]["产品编码"].ToString())
                            {
                                dt.Rows[i]["代理单价"] = plandt.Rows[j].ItemArray[2].ToString();
                            }
                         
                    }
                    }
                }
               
            }
        
            return dt;
        }
        private Microsoft.Office.Interop.Excel.Application m_objExcelApp;             
        private Microsoft.Office.Interop.Excel.Workbook m_objExcelWorkBook;            
        private Microsoft.Office.Interop.Excel.Worksheet m_objExcelWorkSheet;        

        public void  ExportToProvit()
        {

            m_objExcelApp = new Microsoft.Office.Interop.Excel.Application();
            m_objExcelApp.DisplayAlerts = false;
            m_objExcelWorkBook = m_objExcelApp.Workbooks.Open("F:\\cs.xls", Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
            m_objExcelWorkSheet = (Worksheet)m_objExcelWorkBook.Sheets["SAM"];
            PivotCaches objPivot = m_objExcelWorkBook.PivotCaches();
            objPivot.Add(XlPivotTableSourceType.xlDatabase, "SAM!R1C1:R114C11").CreatePivotTable
                ("羽翚!R86C1", "数据透视表1", Type.Missing, Type.Missing);
            m_objExcelWorkBook.PivotTableWizard(Type.Missing, Type.Missing, m_objExcelWorkSheet.Cells[m_objExcelWorkSheet.UsedRange.Rows.Count+3, 1],
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Range objRange = (Range)m_objExcelWorkSheet.Cells[m_objExcelWorkSheet.UsedRange.Rows.Count+3, 1];
            objRange.Select();
            PivotTable objTable = (PivotTable)m_objExcelWorkSheet.PivotTables("数据透视表1");
            PivotField objField = (PivotField)objTable.PivotFields("产品名称");
            objField.Orientation = XlPivotFieldOrientation.xlRowField;
            objField.Position = "1";
            PivotField objFieldN = (PivotField)objTable.PivotFields("数量"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = "1";
            objFieldN = (PivotField)objTable.PivotFields("结算费用"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = "2";
            objFieldN = (PivotField)objTable.PivotFields("补快递费用"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = "3";
            objFieldN = (PivotField)objTable.PivotFields("代理单价"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = "4";
            objTable.DataPivotField.Orientation = XlPivotFieldOrientation.xlColumnField;
            objTable.DataPivotField.Position = "1";

        
            m_objExcelWorkBook.SaveAs("F:\\cs3.xls");
            m_objExcelWorkBook.Close();
            

        }
}
}
