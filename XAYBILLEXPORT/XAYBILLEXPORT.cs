using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kingdee.BOS.Core;
using System.Data;
using Kingdee.BOS.Core.Permission;
using Kingdee.BOS;
using Kingdee.BOS.ServiceHelper;
using Kingdee.BOS.Core.Bill.PlugIn;
using Kingdee.BOS.Core.DynamicForm.PlugIn.Args;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.IO.Pipes;

namespace XAYBILLEXPORT
{
    public class XAYBILLEXPORT : AbstractBillPlugIn
    {
        public override void BarItemClick(BarItemClickEventArgs e)
        {
            base.BarItemClick(e);
            if (e.BarItemKey == "tb_EXPORTTEST")
            {
                string sSQL = "/*dialect*/select a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 编码,floor(b.[FTAXPRICE]) as  拿货单价  ,floor(b.[FTAXAMOUNT]) as 结算费用,(select [FSIMPLENAME] from T_ECC_LOGISTICS_L where FID=wuliu.[FLOGISTICMODE]) as 其他快递, '' as 补快递费用, '' as 补发票费用,'张兰' as 业务员 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FREGIONID]) as 省 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名 ,addr.[FMOBILE] as 电话,wuliu.[FLOGISTICNUM] as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID]";
                DataTable dt = Kingdee.BOS.ServiceHelper.DBServiceHelper.ExecuteDataSet(this.Context, sSQL).Tables[0];
                MemoryStream ms = new MemoryStream();
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, dt);
                byte[] tableBT = ms.ToArray();
                using (NamedPipeServerStream pipeStream = new NamedPipeServerStream("messagepipe", PipeDirection.InOut, 1,PipeTransmissionMode.Message, PipeOptions.None))
                {
                    pipeStream.WaitForConnection();              
                    pipeStream.Write(tableBT, 0, tableBT.Count());     
                                  

                }

            }
        }
        //private void SendData()
        //{
        //    try
        //    {
        //        string sSQL = "/*dialect*/select a.[FBILLNO],a.[FNETORDERDATE],c.[FNAME],b.[FQTY],d.[FNUMBER],b.[FTAXPRICE],b.[FTAXAMOUNT],(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FREGIONID]) as FREGION ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCITYID]) as FCITY,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCOUNTYID]) as FCOUNTY,addr.[FFULLADDRESS],addr.[FRECEIVEPERSON],addr.[FMOBILE],wuliu.[FLOGISTICNUM] from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID]";
        //        DataTable dt = Kingdee.BOS.ServiceHelper.DBServiceHelper.ExecuteDataSet(this.Context, sSQL).Tables[0];
        //        MemoryStream ms = new MemoryStream();
        //        BinaryFormatter bf = new BinaryFormatter();
        //        bf.Serialize(ms, dt);
        //        byte[] tableBT = ms.ToArray();
        //        using (NamedPipeClientStream pipeStream = new NamedPipeClientStream("messagepipe"))
        //        {
        //            pipeStream.Connect();
        //            StreamWriter sw = new StreamWriter(pipeStream);
        //            sw.WriteLine(tableBT);
        //            sw.Flush();
        //            Thread.Sleep(1000);
        //            sw.Close();

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

    }
}
//string sSQL = "/*dialect*/select a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 产品编码,floor(b.[FTAXPRICE]) as 拿货单价,floor(b.[FTAXAMOUNT]) as 结算费用,(select[FSIMPLENAME] from T_ECC_LOGISTICS_L where FID = wuliu.[FLOGISTICMODE]) as 物流公司,floor(b.[F_CAN_KDAMOUNT]) as 补快递费用,floor(b.[F_CAN_TAXAMOUNT]) as 补发票费用,a.[F_CAN_JXS] as 业务员,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FREGIONID]) as 省份 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名,addr.[FMOBILE] as 电话,(case when(wuliu.[FLOGISTICNUM] is null or wuliu.[FLOGISTICNUM] = ' ') then a.[F_CAN_WLDH] else wuliu.[FLOGISTICNUM] end ) as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID] where  DateDiff(dd, a.[FNETORDERDATE], getdate())=0 union all select a.[FDATE], c.[FNAME], floor(b.[FQTY]), d.[FNUMBER],floor(e.[FBILLTAXAMOUNT_LC]),floor(e.[FBILLAMOUNT]),null,null,null,com.[FNAME],(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as FREGION ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as FCITY,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as FCOUNTY,addr.[FFULLADDRESS],addr.[FRECEIVEPERSON],addr.[FMOBILE],'自提' from T_SAL_ORDER as a left join T_SAL_ORDERENTRY as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_SAL_ORDERFIN as e on e.[FENTRYID]=a.[FID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[F_XAY_Base1] left join T_BD_CUSTOMER_L as com on com.[FCUSTID]=a.[FCUSTID] where  DateDiff(dd, a.[FDATE], getdate())=0";