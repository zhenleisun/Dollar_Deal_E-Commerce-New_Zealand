<!--#include file="md5nps.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../config.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      npayid=rs("npayid")
      npaypin=rs("npaypin")
      npayurl=rs("npayurl")
      payall=session("payall")
	  if payall="" then
	 Response.Write "<script Language=Javascript>alert('支付金额不正确,请返回!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>
<%


	curdate=now()
	m_orderid=year(curdate)&month(curdate)&day(curdate)&ipayid&hour(curdate)&minute(curdate)&second(curdate)


	m_id		=	npayid
	modate		=	date+time
	m_orderid	=	m_orderid
	m_oamount	=	payall
	m_ocurrency	=	"1"
	m_url		=	npayurl
	'm_url		=	"http://weburl/pay/NpsVirReceive.asp?do=result"
	m_language	=	"1"
	s_name		=	"cnhww"
	s_addr		=	"cnhww"
	s_postcode	=	"cnhww"
	s_tel		=	"cnhww"			
	s_eml		=	"cnhww"
	r_name		=	"cnhww"
	r_addr		=	"HeBei HengWei Co.,Ltd"
	r_postcode	=	"050000"
	r_tel		=	"013931185013"
	r_eml		=	"cnhww@163.com"
	m_ocomment	=	"节假日送货"	
	m_status	=	"0"
	key		=	npaypin
	
	OrderMessage =m_id&m_orderid&m_oamount&m_ocurrency&m_url&m_language&s_postcode&s_tel&s_eml&r_postcode&r_tel&r_eml&modate&key

	
	
	digest = Ucase(trim(md5(OrderMessage)))
	

	
%>

<div align=center>

<form method="post" action="https://payment.nps.cn/VirReceiveMerchantAction.do" target="_self">
      <input Type="hidden" Name="M_ID" value="<%=m_id%>">
	<input Type="hidden" Name="MOrderID" value="<%=m_orderid%>">
	<input Type="hidden" Name="MOAmount" value="<%=m_oamount%>">
	<input Type="hidden" Name="MOCurrency" value="<%=m_ocurrency%>">
	<input Type="hidden" name="M_URL" value="<%=m_url%>">
	<input Type="hidden" Name="M_Language" value="<%=m_language%>">
	<input Type="hidden" Name="S_Name" value="<%=s_name%>">
	<input Type="hidden" Name="S_Address" value="<%=s_addr%>">
    <input Type="hidden" Name="S_PostCode" value="<%=s_postcode%>">
	<input Type="hidden" Name="S_Telephone" value="<%=s_tel%>">
	<input Type="hidden" Name="S_Email" value="<%=s_eml%>">
	<input Type="hidden" Name="R_Name" value="<%=r_name%>"> 
	<input Type="hidden" Name="R_Address" value="<%=r_addr%>">
	<input Type="hidden" Name="R_PostCode" value="<%=r_postcode%>">
	<input Type="hidden" Name="R_Telephone" value="<%=r_tel%>">
	<input Type="hidden" Name="R_Email" value="<%=r_eml%>">
	<input Type="hidden" name="MOComment" value="<%=m_ocomment%>">
	<input Type="hidden" Name="MODate" value="<%=modate%>">
	<input Type="hidden" Name="State" value="<%=m_status%>">
	<input Type="hidden" Name="digestinfo" value="<%=digest%>">
	<input Type="submit" Name="submit" value="立即使用NPS在线支付"> 
</form>

</div>