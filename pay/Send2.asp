<!--#include file="conn.asp"-->


<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      MerchantID=rs("westid")
      MerchantPostBackURL=rs("westurl")
      MerchantOrderAmount=session("payall")
	  if MerchantOrderAmount="" then
	 Response.Write "<script Language=Javascript>alert('֧������ȷ,�뷵��!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>
<%
	
	'����ϵͳʱ�������������ʽ��YYYYMMDD HHMMSS
	yy=year(date)
	mm=right("00"&month(date),2)
	dd=right("00"&day(date),2)
	ymd=yy&mm&dd

	'���ɶ�������������Ԫ��,��ʽΪ��Сʱ�����ӣ���
	hh=right("00"&hour(time),2)
	nn=right("00"&minute(time),2)
	ss=right("00"&second(time),2)
	hns=hh&nn&ss
	'����8+6=14λ������MerchantOrderNumber
	MerchantOrderNumber=ymd&hns '������
'option explicit
%>
<html>
<body background="../images/bj_x.gif">
<%
''''''''''''''''''''''''''''''''''''''''''
' ��ʾ�������Ͷ���������֧��
''''''''''''''''''''''''''''''''''''''''''
MerchantID=""&MerchantID&""							'�̻�������֧���ϵ�ID
MerchantPostBackURL = ""&MerchantPostBackURL&"/receive.asp?id=" & request.cookies("web767")("username") & ""	'�̻���������֧����֧��֪ͨ����ַ	
MerchantOrderNumber = ""&MerchantOrderNumber&""		'�����ţ������ظ��������á������ڡ����������š�
MerchantOrderAmount = ""&MerchantOrderAmount&""				'����������0.01 ���� 0.01Ԫ

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
        <input type="submit" name="submit" value="�ύ������֧��">
      </form></td>
  </tr>
</table>
<table border="0" width="100%">
  <tr> 
    <td align="center">&nbsp;</td>
  </tr>
</table>
