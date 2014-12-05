<!--#include file="conn.asp"-->


<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      MerchantID=rs("westid")
      MerchantPostBackURL=rs("westurl")
      MerchantOrderAmount=session("payall")
	  if MerchantOrderAmount="" then
	 Response.Write "<script Language=Javascript>alert('支付金额不正确,请返回!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>
<%
	
	'根据系统时间产生订单，格式：YYYYMMDD HHMMSS
	yy=year(date)
	mm=right("00"&month(date),2)
	dd=right("00"&day(date),2)
	ymd=yy&mm&dd

	'生成订单号所有所需元素,格式为：小时，分钟，秒
	hh=right("00"&hour(time),2)
	nn=right("00"&minute(time),2)
	ss=right("00"&second(time),2)
	hns=hh&nn&ss
	'产生8+6=14位订单号MerchantOrderNumber
	MerchantOrderNumber=ymd&hns '订单号
'option explicit
%>
<html>
<body background="../images/bj_x.gif">
<%
''''''''''''''''''''''''''''''''''''''''''
' 该示范程序发送订单到西部支付
''''''''''''''''''''''''''''''''''''''''''
MerchantID=""&MerchantID&""							'商户在西部支付上的ID
MerchantPostBackURL = ""&MerchantPostBackURL&"/receive.asp?id=" & request.cookies("web767")("username") & ""	'商户接收西部支付的支付通知的网址	
MerchantOrderNumber = ""&MerchantOrderNumber&""		'订单号，不能重复。建议用　“日期”＋“订单号”
MerchantOrderAmount = ""&MerchantOrderAmount&""				'订单金额，例：0.01 代表 0.01元

%>
<br>
<center>
</center>
<table border="0" width="100%" bgcolor="#f0f0f0">
  
  <tr align="center"> 
    <td colspan="2"> <form action="http://www.WestPay.com.cn/Pay/WestPayReceiveOrderFromMerchant.asp" method="POST" name="SendOrderToWestPay" target="_self">
        <input type="hidden" name="MerchantID" value="<%=MerchantID%>">
        <input type="hidden" name="PostBackURL" value="<%=MerchantPostBackURL%>">
        <input type="hidden" name="OrderNumber" value="<%=MerchantOrderNumber%>">
        <input type="hidden" name="OrderAmount" value="<%=MerchantOrderAmount%>">
        <input type="submit" name="submit" value="提交到西部支付">
      </form></td>
  </tr>
</table>
<table border="0" width="100%">
  <tr> 
    <td align="center">&nbsp;</td>
  </tr>
</table>
