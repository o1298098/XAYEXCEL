using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace XAYEXCELS
{
    class runreplace
    {
        public string productreplace( string id,string name)
        {
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            DataTable productdt = Form2.ReadFromXml(sysstr + "XAYXML\\wldt.xml");
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
        public DataTable replacedt(DataTable dt)
        {
            string sysstr = System.AppDomain.CurrentDomain.BaseDirectory;
            string plan;
            DataTable productdt = Form2.ReadFromXml(sysstr + "XAYXML\\DataTable.xml");
           
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sales = dt.Rows[i]["业务员"].ToString();
                DataRow[] dr = productdt.Select("代理名字 like " + sales.Substring(0, 1));
                if(dr!=null)
                {
                    plan = dr[0].ItemArray[6].ToString();
                    DataTable plandt = Form2.ReadFromXml(sysstr + "XAYXML\\DataTable.xml");
                    for (int j=0;j<plandt.Rows.Count;j++)
                    {
                        if()
                    }
                }
               
            }
        
            return dt;
        }
    }
}
