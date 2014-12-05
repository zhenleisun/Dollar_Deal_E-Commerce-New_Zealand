<!--#include file="md5.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
	m_id			=	request("m_id")		
	m_orderid		=	request("m_orderid")
	m_oamount		=	request("m_oamount")
	m_ocurrency		=	request("m_ocurrency")
	m_url			=	request("m_url")	
	m_language		=	request("m_language")
	s_name			=	request("s_name")
	s_addr			=	request("s_addr")
	s_postcode		=	request("s_postcode")
	s_tel			=	request("s_tel")
	s_eml			=	request("s_eml")
	r_name			=	request("r_name")
	r_addr			=	request("r_addr")
	r_postcode		=	request("r_postcode")
	r_eml			=	request("r_eml")
	r_tel			=	request("r_tel")
	m_ocomment		=	request("m_ocomment")
	m_status		=	request("m_status")
	modate			=	request("modate")
	newmd5info		=	request("newmd5info")
	key				=	"nps&!@#$"

	if request("newmd5info")="" then
		response.Write("认证签名为空值")
		response.end
	end if

	'升级的
	newOrderMessage = m_id&m_orderid&m_oamount&key&m_status
	
	newMd5text = trim(md5(newOrderMessage))		

	if Ucase(newMd5text)<>Ucase(newmd5info) then
			response.write("认证失败!!!")
	else
			if m_status = 2 then
				response.write	("认证成功，可以更新数据库!!!")		&	"<br>"
				Response.Write "商家号:"				&	m_id		&	"<br>"
				Response.Write "交易订单号:"		&	m_orderid	&	"<br>"
				Response.Write "交易金额:"		&	m_oamount	&	"<br>"
				Response.Write "交易币种: "		&	m_ocurrency &	"<br>"
				Response.Write "语言:"		&	m_language	&	"<br>"
				Response.Write "消费者姓名:"			&	s_name		&	"<br>"
				Response.Write "消费者地址:"			&	s_addr		&	"<br>"
				Response.Write "消费者邮编:"		&	s_postcode	&	"<br>"
				Response.Write "消费者电话:"			&	s_tel		&	"<br>"
				Response.Write "消费者电邮:"			&	s_eml		&	"<br>"
				Response.Write "收货人姓名:"			&	r_name		&	"<br>"
				Response.Write "收货人地址:"			&	r_addr		&	"<br>"
				Response.Write "收货人邮编:"		&	r_postcode	&	"<br>"
				Response.Write "收货人电话:"			&	r_eml		&	"<br>"
				Response.Write "收货人电邮:"			&	r_tel		&	"<br>"
				Response.Write "备注:"		&	m_ocomment	&	"<br>"
				Response.Write "支付状态:"			&	m_status	&	"<br>"
				Response.Write "交易时间:"			&	modate		&	"<br>"
				Response.Write "交易加密码:"			&	newmd5info		&	"<br>"
			else
				Response.Write "支付失败"
			end if
	end if
%>
