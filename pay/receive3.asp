<!--#include file="md5.asp"-->
<%
'''''''''
 ' @Description: ��Ǯ���ؽӿڷ���
 ' @Copyright (c) �Ϻ���Ǯ��Ϣ�������޹�˾
 ' @version 2.0
'''''''''
	key ="99billKeyForTest"		'''�̻���Կ

	merchant_id = request.querystring("merchant_id")			'''��ȡ�̻����
	orderid =  request.querystring("orderid")		'''��ȡ�������
	amount =  request.querystring("amount")	'''��ȡ�������
	date =  request.querystring("date")		'''��ȡ��������
	succeed =  request.querystringr("succeed")	'''��ȡ���׽��,Y�ɹ�,Nʧ��
	mac =  request.querystring("mac")		'''��ȡ��ȫ���ܴ�
	merchant_param =  request.querystring("merchant_param")		'''��ȡ�̻�˽�в���

	couponid = (String)request.getParameter("couponid")		'''��ȡ�Ż�ȯ����
	couponvalue = (String)request.getParameter("couponvalue") 		'''��ȡ�Ż�ȯ���

	'''���ɼ��ܴ�,ע��˳��
	ScrtStr = "merchant_id=" & merchant_id & "&orderid=" & orderid & "&amount=" & amount & "&date=" & mydate & "&succeed=" & succeed & "&merchant_key=" & key
	mac=md5(ScrtStr) 
		

	 v_result="ʧ��"
	if ucase(mac)=ucase(mymac)   then 
			
			if succeed="Y"   then		'''֧���ɹ�
				
				v_result="�ɹ�"
				'''
				'''#�̻���վ�߼�����#
				'''
			
			else		'''֧��ʧ��

			end if

	else		'''ǩ������

	end if
	

%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>��Ǯ99bill</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" />
	</head>
	
	<body>
		
		<div align="center">
		<table width="259" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" >
			<tr bgcolor="#FFFFFF">
				<td width="68">�������:</td>
			  <td width="182"><%=orderid%></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>�������:</td>
			  <td><%=amount%></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>֧�����:</td>
			  <td><%=v_result%></td>
			</tr>
	  </table>
	</div>

	</body>
</html>