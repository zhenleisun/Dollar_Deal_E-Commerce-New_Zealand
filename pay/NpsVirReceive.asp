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
		response.Write("��֤ǩ��Ϊ��ֵ")
		response.end
	end if

	'������
	newOrderMessage = m_id&m_orderid&m_oamount&key&m_status
	
	newMd5text = trim(md5(newOrderMessage))		

	if Ucase(newMd5text)<>Ucase(newmd5info) then
			response.write("��֤ʧ��!!!")
	else
			if m_status = 2 then
				response.write	("��֤�ɹ������Ը������ݿ�!!!")		&	"<br>"
				Response.Write "�̼Һ�:"				&	m_id		&	"<br>"
				Response.Write "���׶�����:"		&	m_orderid	&	"<br>"
				Response.Write "���׽��:"		&	m_oamount	&	"<br>"
				Response.Write "���ױ���: "		&	m_ocurrency &	"<br>"
				Response.Write "����:"		&	m_language	&	"<br>"
				Response.Write "����������:"			&	s_name		&	"<br>"
				Response.Write "�����ߵ�ַ:"			&	s_addr		&	"<br>"
				Response.Write "�������ʱ�:"		&	s_postcode	&	"<br>"
				Response.Write "�����ߵ绰:"			&	s_tel		&	"<br>"
				Response.Write "�����ߵ���:"			&	s_eml		&	"<br>"
				Response.Write "�ջ�������:"			&	r_name		&	"<br>"
				Response.Write "�ջ��˵�ַ:"			&	r_addr		&	"<br>"
				Response.Write "�ջ����ʱ�:"		&	r_postcode	&	"<br>"
				Response.Write "�ջ��˵绰:"			&	r_eml		&	"<br>"
				Response.Write "�ջ��˵���:"			&	r_tel		&	"<br>"
				Response.Write "��ע:"		&	m_ocomment	&	"<br>"
				Response.Write "֧��״̬:"			&	m_status	&	"<br>"
				Response.Write "����ʱ��:"			&	modate		&	"<br>"
				Response.Write "���׼�����:"			&	newmd5info		&	"<br>"
			else
				Response.Write "֧��ʧ��"
			end if
	end if
%>
