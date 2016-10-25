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
using UDPNATCLIENT;
using UDPCOMMON;
using System.Windows.Forms;

namespace XAYBILLEXPORT
{
    public class XAYBILLEXPORT:AbstractListPlugIn
    {
        public Client _client;
        public UserCollection users;
        private void OnUserChanged(UserCollection user)
        {
            users = user;
        }
        public override void AfterBarItemClick(AfterBarItemClickEventArgs e)
        {

            if (e.BarItemKey.EqualsIgnoreCase("tb_EXPORTTEST"))
            {
                NewMethod();
                string sSQL = "/*dialect*/select top 5000 a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 产品编码,floor(b.[FTAXPRICE]) as 拿货单价,floor(b.[FTAXAMOUNT]) as 结算费用,(select[FSIMPLENAME] from T_ECC_LOGISTICS_L where FID = wuliu.[FLOGISTICMODE]) as 物流公司,floor(b.[F_CAN_KDAMOUNT]) as 补快递费用,floor(b.[F_CAN_TAXAMOUNT]) as 补发票费用,a.[F_CAN_JXS] as 业务员,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FREGIONID]) as 省份 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名,addr.[FMOBILE] as 电话,(case when(wuliu.[FLOGISTICNUM] is null or wuliu.[FLOGISTICNUM] = ' ') then a.[F_CAN_WLDH] else wuliu.[FLOGISTICNUM] end ) as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID] where a.[FNETORDERDATE]='2016-07-08'";
                DataTable dt = DBServiceHelper.ExecuteDataSet(this.Context, sSQL).Tables[0];
                MemoryStream ms = new MemoryStream();
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(ms, dt);
                byte[] tableBT = ms.ToArray();
                P2P_TalkMessage msg = new P2P_TalkMessage(tableBT);
                _client.SendMessage(msg, users[0]);
                this.View.ShowMessage("操作成功", Kingdee.BOS.Core.DynamicForm.MessageBoxType.Notice);

            }

        }

        private void NewMethod()
        {
            _client = new Client { OnUserChanged = OnUserChanged };
            _client.Login("K3CLOUD", "");
            _client.Start();
            if (_client != null)
            {
                users = _client._userList;
                Thread.Sleep(1000);
                _client.HolePunching(users[0]);
                
            }
            else
            {
                this.View.ShowMessage("服务器连接失败", Kingdee.BOS.Core.DynamicForm.MessageBoxType.Notice);
            }
        }
    }
}
//string sSQL = "/*dialect*/select a.[FNETORDERDATE] as 日期,c.[FNAME] as 产品名称,floor(b.[FQTY]) as 数量,d.[FNUMBER] as 产品编码,floor(b.[FTAXPRICE]) as 拿货单价,floor(b.[FTAXAMOUNT]) as 结算费用,(select[FSIMPLENAME] from T_ECC_LOGISTICS_L where FID = wuliu.[FLOGISTICMODE]) as 物流公司,floor(b.[F_CAN_KDAMOUNT]) as 补快递费用,floor(b.[F_CAN_TAXAMOUNT]) as 补发票费用,a.[F_CAN_JXS] as 业务员,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FREGIONID]) as 省份 ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCITYID]) as 城市,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where [FDETAILID] = addr.[FCOUNTYID]) as 县区,addr.[FFULLADDRESS] as 地址,addr.[FRECEIVEPERSON] as 姓名,addr.[FMOBILE] as 电话,(case when(wuliu.[FLOGISTICNUM] is null or wuliu.[FLOGISTICNUM] = ' ') then a.[F_CAN_WLDH] else wuliu.[FLOGISTICNUM] end ) as 物流单号 from T_ECC_NETORDER as a left join T_ECC_NETORDEREntry as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[FRECEIVERADDRESS] left join T_ECC_NetOrderLogisticEntry as wuliu on wuliu.[FID]=a.[FID] left join T_ECC_Logistics as Logistics on Logistics.[FID]=a.[FID] where  DateDiff(dd, a.[FNETORDERDATE], getdate())=0 union all select a.[FDATE], c.[FNAME], floor(b.[FQTY]), d.[FNUMBER],floor(e.[FBILLTAXAMOUNT_LC]),floor(e.[FBILLAMOUNT]),null,null,null,com.[FNAME],(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FREGIONID]) as FREGION ,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCITYID]) as FCITY,(select[FDATAVALUE] from T_ECC_LOGISTICSAREADETAIL_L where[FDETAILID] = addr.[FCOUNTYID]) as FCOUNTY,addr.[FFULLADDRESS],addr.[FRECEIVEPERSON],addr.[FMOBILE],'自提' from T_SAL_ORDER as a left join T_SAL_ORDERENTRY as b on b.[FID]= a.[FID] left join T_BD_MATERIAL_L as c on c.[FMATERIALID]=b.[FMATERIALID] left join T_SAL_ORDERFIN as e on e.[FENTRYID]=a.[FID] left join T_BD_MATERIAL as d on d.[FMATERIALID]=b.[FMATERIALID] left join T_ECC_ADDRESS as addr on addr.[FID]=a.[F_XAY_Base1] left join T_BD_CUSTOMER_L as com on com.[FCUSTID]=a.[FCUSTID] where  DateDiff(dd, a.[FDATE], getdate())=0";