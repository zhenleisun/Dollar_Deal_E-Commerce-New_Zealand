<!--#include file="MD5.asp"-->
<!--#include file="conn.asp"-->


<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      payid=rs("payid")
      paypin=rs("paypin")
      payurl=rs("payurl")
      payall=session("payall")
	  if payall="" then
	 Response.Write "<script Language=Javascript>alert('支付金额不正确,请返回!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>
<%
    
	 ' 表单的各项参数如下：

	 ' v_mid
	 '      商户号,这里为测试商户号1001，替换为自己的商户号即可
	 ' key
	 '      MD5私钥
	 ' v_oid
	 '      订单号，构成格式 年月日-商户号-小时分钟秒
	 ' v_amount
	 '      订单金额
	 ' v_moneytype
	 '      支付币种0为人民币
	 ' v_url
	 '      商户自定义返回接收支付结果的页面	 
	 ' remark1 
	 '      备注字段1
	 ' remark2 
	 '      备注字段2
	 ' style
	 '      指网关模式0(普通)，1(银行列表中带外卡)


	 '********以下几项与网上支付货款无关，建议不用**************
	 ' v_rcvname 
	 '      收货人
	 ' v_rcvaddr
	 '      收货地址	       
	 ' v_rcvtel
	 '      订货人电话	
	 ' v_rcvpost
	 '      邮编
	 ' v_ordername
	 '      发货人
	 ' v_orderemail
	 '      订货人EMAIL


	key = paypin
	v_mid = payid
	v_amount=payall
	v_moneytype = "0"
	style="0"
	v_url=payurl
	remark1=""
	remark2=""

	'根据系统时间产生订单，格式：YYYYMMDD-v_mid-HMMSS
	curdate=now()
	v_oid=year(curdate)&month(curdate)&day(curdate)&"-"&v_mid&"-"&hour(curdate)&minute(curdate)&second(curdate)
	text = v_amount&v_moneytype&v_oid&v_mid&v_url&key
	v_md5info=Ucase(trim(md5(text)))					'网银支付平台对MD5值只认大写字符串，所以小写的MD5值得转换为大写

%>

<!--表单确认信息如下-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Payment by DPS</title>
</head>

<body>

<br>
<center>

<body onLoad="javascript:document.E_FORM.submit()">


    <form method="post" action="pxaccess.asp" name="E_FORM" target="_self">

	<input type="hidden" name="v_md5info" size="100"  value="<%=v_md5info%>">
	<input type="hidden" name="v_mid" value="<%=v_mid%>">
	<input type="hidden" name="v_oid" value="<%=v_oid%>">
    <input type="hidden" name="v_amount" value="<%=v_amount%>">
	<input type="hidden" name="v_moneytype"  value="<%=v_moneytype%>">
	<input type="hidden" name="v_url" value="<%=v_url%>">
	<input type="hidden" name="style" value="<%=style%>">
	<input type="hidden" name="remark1" value="<%=remark1%>">
	<input type="hidden" name="remark2" value="<%=remark2%>">

    <input type="hidden" name="TxnType" value="Purchase"><br>
    <input type="hidden" name="AmountInput" value="<%=v_amount%>"><br>
    <input type="hidden" name="MerchantRef" value="post test"><br>

    </form>    


</center>
</body>
</html>