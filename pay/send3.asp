<!--#include file="conn.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="md53.asp"-->

<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      ipayid=rs("ipayid")
      ipaypin=rs("ipaypin")
      ipayurl=rs("ipayurl")
      payall=session("payall")
	  if payall="" then
	 Response.Write "<script Language=Javascript>alert('支付金额不正确,请返回!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>

<%
'''''''''
 ' @Description: 快钱网关接口范例
 ' @Copyright (c) 上海快钱信息服务有限公司
 ' @version 2.0
'''''''''


'根据系统时间产生订单，格式：YYYYMMDD-v_mid-HMMSS
	curdate=now()
	v_oid=year(curdate)&month(curdate)&day(curdate)&ipayid&hour(curdate)&minute(curdate)&second(curdate)
 
	


 merchant_id = ipayid		'''商户编号
 key = ipaypin		'''商户密钥
 orderid = v_oid		'''订单编号
 amount = payall		'''订单金额
 
 tcurrency = "1"	                '''货币类型,1为人民币
 isSupportDES = "2"		'''是否安全校验,2为必校验,推荐

 merchant_url = ipayurl		'''支付结果返回地址
 pname = session("userzhenshiname")		'''支付人姓名
 commodity_info = weburl 		'''商品信息
 merchant_param = ""		'''商户私有参数
	
pemail=""		'''传递email到快钱网关页面订货人的email
pid_99billaccount=""		'''代理/合作伙伴商户编号

'''生成加密串,注意顺序
ScrtStr="merchant_id=" & merchant_id & "&orderid=" & orderid & "&amount=" & amount & "&merchant_url=" & merchant_url & "&merchant_key=" & key
mac=ucase(md5(ScrtStr))

%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>快钱99bill</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" >
	</head>
	
<BODY>
	


		<div align="center" style="font-size=12px;font-weight: bold;color=red;">
		<form name="frm" method="post" action="https://www.99bill.com/webapp/receiveMerchantInfoAction.do">
			<input name="merchant_id" type="hidden" value="<%=merchant_id%>">
			<input name="orderid"  type="hidden" value="<%=orderid%>">
			<input name="amount"  type="hidden" value="<%=amount%>">
			<input name="currency"  type="hidden" value="<%=tcurrency%>">
			<input name="isSupportDES"  type="hidden" value="<%=isSupportDES%>">
			<input name="mac"  type="hidden" value="<%=mac%>">
			
			<input name="merchant_url"  type="hidden"  value="<%=merchant_url%>">
			<input name="pname"  type="hidden" value="<%=pname%>">
			<input name="commodity_info"  type="hidden"  value="<%=commodity_info%>">
			<input name="merchant_param" type="hidden"  value="<%=merchant_param%>">

			<input name="pemail" type="hidden"  value="<%=pemail%>">
			<input name="pid_99billaccount" type="hidden"  value="<%=pid_99billaccount%>">
			
				
    <input type=submit name=v_action value="立即使用快钱支付网关支付">
		</form>
          
</div>

</BODY>
</HTML>