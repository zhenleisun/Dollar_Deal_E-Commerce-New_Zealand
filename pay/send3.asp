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
	 Response.Write "<script Language=Javascript>alert('֧������ȷ,�뷵��!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>

<%
'''''''''
 ' @Description: ��Ǯ���ؽӿڷ���
 ' @Copyright (c) �Ϻ���Ǯ��Ϣ�������޹�˾
 ' @version 2.0
'''''''''


'����ϵͳʱ�������������ʽ��YYYYMMDD-v_mid-HMMSS
	curdate=now()
	v_oid=year(curdate)&month(curdate)&day(curdate)&ipayid&hour(curdate)&minute(curdate)&second(curdate)
 
	


 merchant_id = ipayid		'''�̻����
 key = ipaypin		'''�̻���Կ
 orderid = v_oid		'''�������
 amount = payall		'''�������
 
 tcurrency = "1"	                '''��������,1Ϊ�����
 isSupportDES = "2"		'''�Ƿ�ȫУ��,2Ϊ��У��,�Ƽ�

 merchant_url = ipayurl		'''֧��������ص�ַ
 pname = session("userzhenshiname")		'''֧��������
 commodity_info = weburl 		'''��Ʒ��Ϣ
 merchant_param = ""		'''�̻�˽�в���
	
pemail=""		'''����email����Ǯ����ҳ�涩���˵�email
pid_99billaccount=""		'''����/��������̻����

'''���ɼ��ܴ�,ע��˳��
ScrtStr="merchant_id=" & merchant_id & "&orderid=" & orderid & "&amount=" & amount & "&merchant_url=" & merchant_url & "&merchant_key=" & key
mac=ucase(md5(ScrtStr))

%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>��Ǯ99bill</title>
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
			
				
    <input type=submit name=v_action value="����ʹ�ÿ�Ǯ֧������֧��">
		</form>
          
</div>

</BODY>
</HTML>