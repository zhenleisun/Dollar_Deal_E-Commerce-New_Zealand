<!-- #include file="MD5.asp" -->
<%
'*******************************************
'�ļ�����GetPayNotify.asp
'��Ҫ���ܣ���ʾ��������Ҫ��ɽ�������֧������֧��֪ͨ��Ϣ����֤��Ϣ��Ч�ԣ��ж�֧��������������߼���������֪ͨ����
'�汾��v1.6��Build2005-05-24��
'˵����
'	1.��ҳ���벻Ҫʹ������response.redirect��ҳ��ת������
'	2.��ֱ�ӽ�֪ͨ���������XML����ʽ����ڱ�ҳ������֧��@�����������������
'	3.������ڱ�ҳ����д����̻��Լ���������򣬽��з�����ϵ�в���
'��Ȩ���У����������������缼�����޹�˾
'����֧����ϵ��ʽ��86-010-84853768/69 ת������
'*******************************************

'---��ȡ����֧���������̻����͵�֧��֪ͨ��Ϣ(���¼��Ϊ֪ͨ��Ϣ)
	c_mid			= request("c_mid")			'�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
	c_order			= request("c_order")		'�̻��ṩ�Ķ�����
	c_orderamount	= request("c_orderamount")	'�̻��ṩ�Ķ����ܽ���ԪΪ��λ��С���������λ���磺13.05
	c_ymd			= request("c_ymd")			'�̻���������Ķ����������ڣ���ʽΪ"yyyymmdd"����20050102
	c_transnum		= request("c_transnum")		'����֧�������ṩ�ĸñʶ����Ľ�����ˮ�ţ����պ��ѯ���˶�ʹ�ã�
	c_succmark		= request("c_succmark")		'���׳ɹ���־��Y-�ɹ� N-ʧ��			
	c_moneytype		= request("c_moneytype")	'֧�����֣�0Ϊ�����
	c_cause			= request("c_cause")		'�������֧��ʧ�ܣ����ֵ����ʧ��ԭ��		
	c_memo1			= request("c_memo1")		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�����һ
	c_memo2			= request("c_memo2")		'�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�������
	c_signstr		= request("c_signstr")		'����֧�����ض�������Ϣ����MD5���ܺ���ַ���

	'---У����Ϣ������---
		IF c_mid="" or c_order="" or c_orderamount="" or c_ymd="" or c_moneytype="" or c_transnum="" or c_succmark="" or c_signstr="" THEN
			response.write "֧����Ϣ����"
			response.end
		END IF

	'---����õ�֪ͨ��Ϣƴ���ַ�������Ϊ׼������MD5���ܵ�Դ������Ҫע����ǣ���ƴ��ʱ���Ⱥ�˳���ܸı�
		Dim c_pass	'�̻���֧����Կ����¼�̻������̨(https://www.cncard.net/admin/)���ڹ�����ҳ���ҵ���ֵ
		c_pass = "Test"
		
		srcStr = c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & c_pass

	'---��֧��֪ͨ��Ϣ����MD5����
		r_signstr	= MD5(srcStr)

	'---У���̻���վ��֪ͨ��Ϣ��MD5���ܵĽ��������֧�������ṩ��MD5���ܽ���Ƿ�һ��
		IF r_signstr<>c_signstr THEN
			response.write "ǩ����֤ʧ��"
			response.end
		END IF

	'---У���̻����
		Dim MerchantID	'�̻��Լ��ı��
		IF MerchantID<>c_mid THEN
			response.write "�ύ���̻��������"
			response.end
		END IF

	'---У���̻�����ϵͳ���Ƿ���֪ͨ��Ϣ���صĶ�����Ϣ
		Dim conn	'�̻�ϵͳ����������
		sql="select top 1 ������ from �̻��Ķ����� where �̻�������="& c_order
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn
		IF rs.eof THEN
			response.write "δ�ҵ��ö�����Ϣ"
			response.end
		END IF

	'---У���̻�����ϵͳ�м�¼�Ķ�����������֧������֪ͨ��Ϣ�еĽ���Ƿ�һ��
		Dim r_orderamount	'�̻��Լ�ϵͳ��¼�Ķ������
		r_orderamount=rs("�������")	'�̻����Լ�����ϵͳ��ȡ��ֵ
		IF ccur(r_orderamount)<>ccur(c_orderamount) THEN
			response.write "֧���������"
			response.end
		END IF

	'---У���̻�����ϵͳ�м�¼�Ķ����������ں�����֧������֪ͨ��Ϣ�еĶ������������Ƿ�һ��
		Dim r_ymd	'�̻��Լ�ϵͳ��¼�Ķ�����������
		r_ymd=rs("������������")	'�̻����Լ�����ϵͳ��ȡ��ֵ
		IF r_ymd<>c_ymd THEN
			response.write "����ʱ������"
			response.end
		END IF

	'---У���̻�ϵͳ�м�¼����Ҫ��֧�����֪ͨ��ת���Ĳ���������֧������֪ͨ��Ϣ���ṩ�Ĳ����Ƿ�һ��
		Dim r_memo1	'�̻��Լ�ϵͳ��¼����Ҫ��֧�����֪ͨ��ת���Ĳ���һ
		r_memo1 = rs("ת������һ")
		Dim r_memo2	'�̻��Լ�ϵͳ��¼����Ҫ��֧�����֪ͨ��ת���Ĳζ�
		r_memo2 = rs("ת��������")
		IF r_memo1<>c_memo1 or r_memo2<>c_memo2 THEN
			response.write "�����ύ����"
			response.end
		END IF

	'---У�鷵�ص�֧������ĸ�ʽ�Ƿ���ȷ
		IF c_succmark<>"Y" and c_succmark<>"N" THEN
			response.write "�����ύ����"
			response.end
		END IF

	'---���ݷ��ص�֧��������̻������Լ��ķ����Ȳ���
		IF c_succmark="Y" THEN
			'�����̻��Լ�������򣬽��з�����ϵ�в���
		END IF

	'---���֧�����֪ͨ����
		'<result>��ֵ�̶�Ϊ1����ʾ�̻��ѳɹ��յ����ص�֧���ɹ���֪ͨ��
		'<reURL> ���̻���ʾ���û�������ҳ���URL(��Ӧ�����ļ���GetPayHandle.asp)
		Response.Write("<result>1</result><reURL>http://www.yoursitename.com/urlpath/GetPayHandle.asp</reURL>")
		Response.End()
%>