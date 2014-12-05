<!--#include file="md5.asp"-->
<%
'''''''''
 ' @Description: 快钱网关接口范例
 ' @Copyright (c) 上海快钱信息服务有限公司
 ' @version 2.0
'''''''''
	key ="99billKeyForTest"		'''商户密钥

	merchant_id = request.querystring("merchant_id")			'''获取商户编号
	orderid =  request.querystring("orderid")		'''获取订单编号
	amount =  request.querystring("amount")	'''获取订单金额
	date =  request.querystring("date")		'''获取交易日期
	succeed =  request.querystringr("succeed")	'''获取交易结果,Y成功,N失败
	mac =  request.querystring("mac")		'''获取安全加密串
	merchant_param =  request.querystring("merchant_param")		'''获取商户私有参数

	couponid = (String)request.getParameter("couponid")		'''获取优惠券编码
	couponvalue = (String)request.getParameter("couponvalue") 		'''获取优惠券面额

	'''生成加密串,注意顺序
	ScrtStr = "merchant_id=" & merchant_id & "&orderid=" & orderid & "&amount=" & amount & "&date=" & mydate & "&succeed=" & succeed & "&merchant_key=" & key
	mac=md5(ScrtStr) 
		

	 v_result="失败"
	if ucase(mac)=ucase(mymac)   then 
			
			if succeed="Y"   then		'''支付成功
				
				v_result="成功"
				'''
				'''#商户网站逻辑处理#
				'''
			
			else		'''支付失败

			end if

	else		'''签名错误

	end if
	

%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>快钱99bill</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" />
	</head>
	
	<body>
		
		<div align="center">
		<table width="259" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" >
			<tr bgcolor="#FFFFFF">
				<td width="68">订单编号:</td>
			  <td width="182"><%=orderid%></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>订单金额:</td>
			  <td><%=amount%></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>支付结果:</td>
			  <td><%=v_result%></td>
			</tr>
	  </table>
	</div>

	</body>
</html>