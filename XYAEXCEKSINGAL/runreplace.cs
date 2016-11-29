﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace XYAEXCEKSINGAL
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
                if (sales!="") {
                    DataRow[] dr = productdt.Select("代理名字 like '" + sales.Substring(0, 1) + "%'");
                    if (dr.Length > 0)
                    {
                        plan = dr[0].ItemArray[6].ToString();
                        if (plan != "")
                        {
                            System.Data.DataTable plandt = Form2.ReadFromXml(sysstr + "XAYXML\\" + plan + ".xml");
                            for (int j = 0; j < plandt.Rows.Count; j++)
                            {
                                if (plandt.Rows[j].ItemArray[1].ToString() == dt.Rows[i]["产品编码"].ToString())
                                {
                                    dt.Rows[i]["拿货单价"] = Convert.ToInt32(plandt.Rows[j].ItemArray[2].ToString() == "" ? "0" : plandt.Rows[j].ItemArray[2].ToString());
                                    dt.Rows[i]["结算费用"] = Convert.ToInt32(plandt.Rows[j].ItemArray[2].ToString() == "" ? "0" : plandt.Rows[j].ItemArray[2].ToString()) * Convert.ToInt32(dt.Rows[i].ItemArray[2].ToString());
                                }

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

        public void  ExportToProvit(string filename,bool KD,bool FP)
        {
           
                m_objExcelApp = new Microsoft.Office.Interop.Excel.Application();
                m_objExcelApp.DisplayAlerts = false;
                m_objExcelWorkBook = m_objExcelApp.Workbooks.Open(filename, Type.Missing,
                                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
                int sb = m_objExcelApp.Workbooks[1].Worksheets.Count;
                string[] sheetName = new string[50];
                string str1;
                for (int k = 1; k < sb ; k++)
                {
                sheetName [k-1]= m_objExcelApp.Worksheets[k + 1].Name;
               
            }
            for (int k = 1; k < sb ; k++)
            {
                str1 = sheetName[k - 1];

                NewMethod(str1,KD,FP);
            
            }



            m_objExcelWorkBook.SaveAs(filename);
                m_objExcelWorkBook.Close();
          

        }

        private void NewMethod(string sheetName, bool KD, bool FP)
        {
            m_objExcelWorkSheet = (Worksheet)m_objExcelWorkBook.Sheets[sheetName];
            m_objExcelWorkSheet.Activate();
            PivotCaches objPivot = m_objExcelWorkBook.PivotCaches();
            string rangedata = sheetName + "!R1C1:R" + m_objExcelWorkSheet.UsedRange.Rows.Count + "C11";
            objPivot.Add(XlPivotTableSourceType.xlDatabase,rangedata).CreatePivotTable
                (m_objExcelWorkSheet.Cells[m_objExcelWorkSheet.UsedRange.Rows.Count + 3, 1], sheetName+"1", Type.Missing, XlPivotTableVersionList.xlPivotTableVersion15);
            Range objRange = (Range)m_objExcelWorkSheet.Cells[m_objExcelWorkSheet.UsedRange.Rows.Count + 3, 1];
            objRange.Select();
            PivotTable objTable = (PivotTable)m_objExcelWorkSheet.PivotTables(sheetName+"1");
            PivotField objField = (PivotField)objTable.PivotFields("产品名称");
            objField.Orientation = XlPivotFieldOrientation.xlRowField;
            objField.Position = "1";
            int count = 1;
            PivotField objFieldN = (PivotField)objTable.PivotFields("产品数量"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = count;
            count++;
            objFieldN = (PivotField)objTable.PivotFields("结算费用"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = count;
            count++;
            if(!KD)
            {
            objFieldN = (PivotField)objTable.PivotFields("补快递费用"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = count;
            count++;
            }
            if (!FP) {
            objFieldN = (PivotField)objTable.PivotFields("补发票费用"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = count;
            count++;
            }
            objFieldN = (PivotField)objTable.PivotFields("代理价"); //赋值数据区域数据
            objFieldN.Orientation = XlPivotFieldOrientation.xlDataField;
            objFieldN.Position = count;
            objTable.DataPivotField.Orientation = XlPivotFieldOrientation.xlColumnField;
            objTable.DataPivotField.Position = "1";
            
        }
    }
}