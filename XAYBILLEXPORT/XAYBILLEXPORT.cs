using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using Kingdee.BOS.ServiceHelper;
using Kingdee.BOS.Core.DynamicForm.PlugIn.Args;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Kingdee.BOS.Core.List.PlugIn;
using Kingdee.BOS.Util;
using System.Threading;
using P2PCLIENT;
using System.Text;

namespace XAYBILLEXPORT
{
    public class XAYBILLEXPORT : AbstractListPlugIn
    {

        string removeip = "";
        public override void AfterBarItemClick(AfterBarItemClickEventArgs e)
        {
           
            if (e.BarItemKey.EqualsIgnoreCase("tb_EXPORTTEST"))
            {
                ClientInfo MClient = new ClientInfo("120.76.230.35", 56422, 4000, 1000, 60000, 2, "00133E02D81B");
                MClient.ClientConnToMe += new ClientConnToHandler(client_ClientConnToMe);
                MClient.ConToServer();
                string sSQL = "select a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 产品编码,floor(b.[FTAXPRICE]) as 拿货单价,floor(b.[FTAXAMOUNT]) as 结算费用,(select[FSIMPLENAME] from T_ECC_LOGISTICS_L where FID=wuliu.[FLOGISTICMODE]) as 物流公司,floor(b.[F_CAN_KDAMOUNT]) as 补快递费用,floor(b.[F_CAN_TAXAMOUNT]) as 补发票费用,a.[F_CAN_JXS] as 业务员,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as 省份 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名,addr.[FMOBILE] as 电话,(case when(wuliu.[FLOGISTICNUM] is null or wuliu.[FLOGISTICNUM]= ' ') then a.[F_CAN_WLDH] else wuliu.[FLOGISTICNUM] end) as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID] where  DateDiff(dd, a.[FNETORDERDATE], getdate())=0 union all select a.[FDATE], c.[FNAME],floor(b.[FQTY]),d.[FNUMBER],floor(e.[FTAXPRICE]),floor(e.[FALLAMOUNT]),null,null,null,com.[FNAME],(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as FREGION ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as FCITY,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as FCOUNTY,addr.[FFULLADDRESS],addr.[FRECEIVEPERSON],addr.[FMOBILE],'自提' from T_SAL_ORDER as a left join T_SAL_ORDERENTRY as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_SAL_ORDERENTRY_F as e on e.[FID]=a.[FID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[F_XAY_Base1] left join T_BD_CUSTOMER_L as com on com.[FCUSTID]=a.[FCUSTID] where  DateDiff(dd, a.[FDATE], getdate())=0 and a.[FBILLTYPEID]= 'eacb50844fc84a10b03d7b841f3a6278'";
                DataTable dt = DBServiceHelper.ExecuteDataSet(this.Context, sSQL).Tables[0];
                MemoryStream ms = new MemoryStream();
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, dt);
                byte[] tableBT = ms.ToArray();
                Thread.Sleep(4000);
                List<string> userlist = MClient.GetAllUser();
                foreach (var item in userlist)
                {
                        MClient.SendData(item, tableBT);
                        byte[] endByte = Encoding.Default.GetBytes("☆");
                        MClient.SendData(item, endByte);
                }

                MClient.Mainclient.Sock.Close();
                this.View.ShowMessage("操作成功", Kingdee.BOS.Core.DynamicForm.MessageBoxType.Notice);



            }

          
            }
        void client_ClientConnToMe(ConClient client)
        {
            if (client.Host.Contains(this.Context.ClientInfo.IpAddress))
            {
                removeip = client.Key;
            }
        }

    }
}
//string sSQL = "select a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 产品编码,floor(b.[FTAXPRICE]) as 拿货单价,floor(b.[FTAXAMOUNT]) as 结算费用,(select[FSIMPLENAME] from T_ECC_LOGISTICS_L where FID=wuliu.[FLOGISTICMODE]) as 物流公司,floor(b.[F_CAN_KDAMOUNT]) as 补快递费用,floor(b.[F_CAN_TAXAMOUNT]) as 补发票费用,a.[F_CAN_JXS] as 业务员,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as 省份 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名,addr.[FMOBILE] as 电话,(case when(wuliu.[FLOGISTICNUM] is null or wuliu.[FLOGISTICNUM]= ' ') then a.[F_CAN_WLDH] else wuliu.[FLOGISTICNUM] end) as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID] /*where  DateDiff(dd, a.[FNETORDERDATE], getdate())=0*/ where a.[FNETORDERDATE]='2016-10-26'union all select a.[FDATE], c.[FNAME],floor(b.[FQTY]),d.[FNUMBER],floor(e.[FTAXPRICE]),floor(e.[FALLAMOUNT]),null,null,null,com.[FNAME],(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as FREGION ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as FCITY,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as FCOUNTY,addr.[FFULLADDRESS],addr.[FRECEIVEPERSON],addr.[FMOBILE],'自提' from T_SAL_ORDER as a left join T_SAL_ORDERENTRY as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_SAL_ORDERENTRY_F as e on e.[FID]=a.[FID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[F_XAY_Base1] left join T_BD_CUSTOMER_L as com on com.[FCUSTID]=a.[FCUSTID]  /*where  DateDiff(dd, a.[FDATE], getdate())=0*/  where a.[FDATE]='2016-10-26' and a.[FBILLTYPEID]= 'eacb50844fc84a10b03d7b841f3a6278';";