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
	 Response.Write "<script Language=Javascript>alert('֧������ȷ,�뷵��!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>
<%
    
	 ' ���ĸ���������£�

	 ' v_mid
	 '      �̻���,����Ϊ�����̻���1001���滻Ϊ�Լ����̻��ż���
	 ' key
	 '      MD5˽Կ
	 ' v_oid
	 '      �����ţ����ɸ�ʽ ������-�̻���-Сʱ������
	 ' v_amount
	 '      �������
	 ' v_moneytype
	 '      ֧������0Ϊ�����
	 ' v_url
	 '      �̻��Զ��巵�ؽ���֧�������ҳ��	 
	 ' remark1 
	 '      ��ע�ֶ�1
	 ' remark2 
	 '      ��ע�ֶ�2
	 ' style
	 '      ָ����ģʽ0(��ͨ)��1(�����б��д��⿨)


	 '********���¼���������֧�������޹أ����鲻��**************
	 ' v_rcvname 
	 '      �ջ���
	 ' v_rcvaddr
	 '      �ջ���ַ	       
	 ' v_rcvtel
	 '      �����˵绰	
	 ' v_rcvpost
	 '      �ʱ�
	 ' v_ordername
	 '      ������
	 ' v_orderemail
	 '      ������EMAIL


	key = paypin
	v_mid = payid
	v_amount=payall
	v_moneytype = "0"
	style="0"
	v_url=payurl
	remark1=""
	remark2=""

	'����ϵͳʱ�������������ʽ��YYYYMMDD-v_mid-HMMSS
	curdate=now()
	v_oid=year(curdate)&month(curdate)&day(curdate)&"-"&v_mid&"-"&hour(curdate)&minute(curdate)&second(curdate)
	text = v_amount&v_moneytype&v_oid&v_mid&v_url&key
	v_md5info=Ucase(trim(md5(text)))					'����֧��ƽ̨��MD5ֵֻ�ϴ�д�ַ���������Сд��MD5ֵ��ת��Ϊ��д

%>

<!--��ȷ����Ϣ����-->
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