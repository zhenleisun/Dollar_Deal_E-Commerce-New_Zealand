
<%
set rsb=server.CreateObject("adodb.recordset")
rsb.open "select * from webinfo ",conn,1,3


	seller_email	= rsb("alipayid")	 '����дǩԼ֧�����˺ţ�
	partner			= rsb("partner")	 '��дǩԼ֧�����˺Ŷ�Ӧ��partnerID��
	key			    = rsb("alipaymd5")	 '��дǩԼ�˺Ŷ�Ӧ�İ�ȫУ����

    notify_url		= "./Alipay_Notify.asp"	        '���׹����з�����֪ͨ��ҳ�� Ҫ�� http://��ʽ������·��������http://www.alipay.com/alipay/Alipay_Notify.asp  ע���ļ�λ������д��ȷ��
	return_url		= "./return_Alipay_Notify.asp"	'��������ת��ҳ�� Ҫ�� http://��ʽ������·��, ����http://www.alipay.com/alipay/return_Alipay_Notify.asp  ע���ļ�λ������д��ȷ��
	'���ʹ����Alipay_Notify.asp����return_Alipay_Notify.asp�������������ļ��������Ӧ�ĺ��������ID�Ͱ�ȫУ����
	logistics_fee	   = "0.00"			'�������ͷ���
	logistics_payment  = "BUYER_PAY"	'�������ͷ��ø��ʽ��SELLER_PAY(����֧��)��BUYER_PAY(���֧��)��BUYER_PAY_AFTER_RECEIVE(��������)
	logistics_type	   = "EXPRESS"		'�������ͷ�ʽ��POST(ƽ��)��EMS(EMS)��EXPRESS(�������)

	 	 
'��½ www.alipay.com ��, ���̼ҷ���,���Կ���֧������ȫУ����ͺ���id,������������ 
%>