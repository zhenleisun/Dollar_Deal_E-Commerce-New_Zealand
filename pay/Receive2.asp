<!--#include file="conn.asp"-->
<%

set rs=server.CreateObject("adodb.recordset")
rs.Open "select payid from webinfo ",conn,1,3
MerchantID_User=rs("westid")'֧��ID(�ͻ�ID)
rs.close
set rs=nothing
''''''''''''''''''''''''''''''''''''''''''
' ��ʾ�������������֧��������֧���ɹ�֪ͨ����֤����ʵ�ԣ�������Ӧ���Զ�����
''''''''''''''''''''''''''''''''''''''''''

dim MerchantOrderNumber, WestPayOrderNumber, PaidAmount, MerchantID

MerchantOrderNumber=request("MerchantOrderNumber") '���̻�֧�������еĶ�������ͬ
WestPayOrderNumber=request("WestPayOrderNumber")
PaidAmount=ccur("0"&request("PaidAmount"))  'WestPay���ص�ʵ��֧������CCURתΪ�����͡�
'ע���̻�����������Ǵ����̻�ԭʼ�������ҵ�ԭʼ�������Ƚ�ʵ������ԭʼ��������ͬ����֧���ɹ���
MerchantID=request("MerchantID")
'ע���̻������жϴ��̻�ID�ǲ��������̻�ID

Dim objHttp, str

Dim PaymentCompleted
PaymentCompleted=false

' ׼���ش�֧��֪ͨ��
str = Request.Form & "&cmd=validate"

'set objHttp = Server.CreateObject("Msxml21.ServerXMLHTTP")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
 
'��WestPay������֪ͨ�����ٴ��ص�WestPay����֤��ȷ��֪ͨ��Ϣ����ʵ�� 
 
objHttp.open "POST", "http://www.WestPay.com.cn/pay/ISPN.asp", false
'ISPN: Instant Secure Payment Notification

objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

if (objHttp.status <> 200 ) then
	    'HTTP ������
		response.Write("Status="&objHttp.status)
elseif (objHttp.responseText = "VERIFIED") then
	
		'֧��֪ͨ��֤�ɹ�
		if MerchantID=""&MerchantID_User&"" then '�жϴ˶����ǲ����̻��Ķ�����
			PaymentCompleted=true
		end if
	
elseif (objHttp.responseText = "INVALID") then
		'֧��֪ͨ��֤ʧ��
		response.Write("Invalid")
else
		'֧��֪ͨ��֤�����г��ִ���
		response.Write("Error")
end if
set objHttp = nothing

if PaymentCompleted then
	'֧���ɹ��Ĵ���...�����³�������ο���

	'ע���̻�����������Ǵ����̻�ԭʼ�������ҵ�ԭʼ�������Ƚ�ʵ������ԭʼ��������ͬ����֧���ɹ���
	'�˴�Ϊ���߳�ֵ,���ý��ȶ��Ƿ�!
	
	'if ����ԭʼ�������=PaidAmount then
	reseponse.write "��ϲ��֧���ɹ�!"
	'else
		'ʧ��
	'end if
else

	'֧�����ɹ��Ĵ���...
response.redirect "logout3.asp"end if
%>�