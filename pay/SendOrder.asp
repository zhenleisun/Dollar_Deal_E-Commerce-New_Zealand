<!-- #include file="YMD5.asp" -->
<!--#include file="conn.asp"-->
<%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      payid=rs("ypayid")
      paypin=rs("ypaypin")
      payurl=rs("ypayurl")
      payall=session("payall")
	  if payall="" then
	 Response.Write "<script Language=Javascript>alert('֧������ȷ,�뷵��!');location.href = 'javascript:history.back()';</script>"
response.end 
end if 
%>

<%
'*******************************************
'�ļ�����SendOrder.asp
'��Ҫ���ܣ���ʾ��������Ҫ��ɽ��̻�������Ϣ�ύ������֧�����صĹ���
'�汾��v1.6��Build2005-05-24��
'�����������̻��Ķ���ϵͳ������ɣ���ҳ����Ҫ�ǰ����̻���������֧������Ҫ��ĸ�ʽ��������Ϣ�ύ������֧��@����֧���ӿڣ�����֧������
'��Ȩ���У����������������缼�����޹�˾
'����֧����ϵ��ʽ��86-010-84853768/69 ת������
'*******************************************

'---������Ϣ---
	Dim c_mid			'�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
	Dim c_order			'�̻���վ���ɵĶ����ţ������ظ�
	Dim c_name			'�̻������е��ջ�������
	Dim c_address		'�̻������е��ջ��˵�ַ
	Dim c_tel			'�̻������е��ջ��˵绰
	Dim c_post			'�̻������е��ջ����ʱ�
	Dim c_email			'�̻������е��ջ���Email
	Dim c_orderamount	'�̻������ܽ��
	Dim c_ymd			'�̻������Ĳ������ڣ���ʽΪ"yyyymmdd"����20050102
	Dim c_moneytype		'֧�����֣�0Ϊ�����
	Dim c_retflag		'�̻�����֧���ɹ����Ƿ���Ҫ�����̻�ָ�����ļ���0�����÷��� 1����Ҫ����
	Dim c_paygate		'������̻���վѡ�����������ø�ֵ������ֵ�ɲμ�������֧��@�������ӿ��ֲᡷ��¼һ�����������֧��@��ѡ�����д���Ϊ��ֵ��
	Dim c_returl		'���c_retflagΪ1ʱ����ֵ����֧���ɹ��󷵻ص��ļ���·��
	Dim c_memo1			'�̻���Ҫ��֧�����֪ͨ��ת�����̻�����һ
	Dim c_memo2			'�̻���Ҫ��֧�����֪ͨ��ת�����̻�������
	Dim c_signstr		'�̻��Զ�����Ϣ����MD5ǩ������ַ���
	Dim c_pass			'֧����Կ�����¼�̻������̨�����ʻ���Ϣ->������Ϣ->��ȫ��Ϣ�е�֧����Կ��
	Dim notifytype		'0��֪ͨͨ��ʽ/1������֪ͨ��ʽ����ֵΪ��֪ͨͨ��ʽ
	Dim c_language		'�������˹��ʿ�֧��ʱ����ʹ�ø�ֵ����������������֧��ʱ��ҳ�����֣�ֵΪ��0����ҳ����ʾΪ����/1����ҳ����ʾΪӢ��
	
	
		curdate=now()
	v_oid=year(curdate)&month(curdate)&day(curdate)


	c_mid		= payid
	c_order		= v_oid
	c_name		= "�̳Ƕ���"
	c_address	= "�����г���������·33��"
	c_tel		= "010-88128382"
	c_post		= "100001"
	c_email		= "stcdver@tat.com"
	c_orderamount	= payall
	c_ymd		= "20060102"
	c_moneytype	= "0"
	c_retflag	= "1"
	c_paygate	= ""
	c_returl	= payurl	'�õ�ַΪ�̻���������֧�����֪ͨ��ҳ�棬���ύ�����ļ���(��Ӧ�����ļ���GetPayNotify.asp)
	c_memo1		= "ABCDE"
	c_memo2		= "12345"
	c_pass		= paypin
	notifytype	= "1"
	c_language	= "0"

	srcStr = c_mid & c_order & c_orderamount & c_ymd & c_moneytype & c_retflag & c_returl & c_paygate & c_memo1 & c_memo2 & notifytype & c_language & c_pass
	'˵�����������ָ��֧����ʽ(c_paygate)��ֵʱ����Ҫ�����û�ѡ��֧����ʽ��Ȼ���ٸ����û�ѡ��Ľ�����������MD5���ܣ�Ҳ����˵����ʱ����ҳ��Ӧ�ò��Ϊ����ҳ�棬��Ϊ����������ɡ�
	
'---�Զ�����Ϣ����MD5����
	c_signstr	= MD5(srcStr)


%>

<html>
<body onLoad="javascript:document.payForm1.submit()">
<p>&nbsp;</p>
<table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"> 
      <form name="payForm1" action="https://www.cncard.net/purchase/getorder.asp" method="POST">
			<input type="hidden" name="c_mid" value="<%=c_mid%>">
			<input type="hidden" name="c_order" value="<%=c_order%>">
			<input type="hidden" name="c_name" value="<%=c_name%>">
			<input type="hidden" name="c_address" value="<%=c_address%>">
			<input type="hidden" name="c_tel" value="<%=c_tel%>">
			<input type="hidden" name="c_post" value="<%=c_post%>">
			<input type="hidden" name="c_email" value="<%=c_email%>">
			<input type="hidden" name="c_orderamount" value="<%=c_orderamount%>">
			<input type="hidden" name="c_ymd" value="<%=c_ymd%>">
			<input type="hidden" name="c_moneytype" value="<%=c_moneytype%>">
			<input type="hidden" name="c_retflag" value="<%=c_retflag%>">
			<input type="hidden" name="c_paygate" value="<%=c_paygate%>">
			<input type="hidden" name="c_returl" value="<%=c_returl%>">
			<input type="hidden" name="c_memo1" value="<%=c_memo1%>">
			<input type="hidden" name="c_memo2" value="<%=c_memo2%>">
			<input type="hidden" name="c_language" value="<%=c_language%>">
			<input type="hidden" name="notifytype" value="<%=notifytype%>">
			<input type="hidden" name="c_signstr" value="<%=c_signstr%>">
       </form>
	</td>
  </tr>
</table>
<p>
<p>&nbsp; </p>
</body>
</html>