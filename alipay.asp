<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="alipayto/alipay_payto.asp"-->
 <%  qian=trim(request("p"))
 
 d=trim(request("d"))

	 shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
'����Ĳ���	
    service         =   "create_partner_trade_by_buyer"   'trade_create_by_buyer ��ʾ��׼˫�ӿڣ� create_partner_trade_by_buyer ��ʾ�������׽ӿ�
	subject			=	d '��Ʒ���ƣ�
	body			=	webname&"����"		'��Ʒ����
	out_trade_no    =   dingdan        '��ʱ���ȡ�Ķ����ţ������޸ĳ��Լ���վ�Ķ����ţ���֤ÿ���ύ�Ķ���Ψһ����
	

     price		= qian	'��Ʒ�۸�

    quantity        =   "1"             '��Ʒ����,����߹��ﳵĬ��Ϊ1
    seller_email    =    seller_email   '���� ��֧�����ʺţ�c2c�ͻ������Ը��Ĵ˲�����

 '�����ǿ�ѡ�Ĳ��� ���û�п���Ϊ�ա�ע�⣺��������ϵ��ַ���������� ������Ҫô��Ϊ�գ�Ҫô������Ϊ�ա�
  	show_url        = ""  '�̻���չʾ��ַ�����Ӻ��治���Զ������
	receive_name    = ""  '�ջ�������
    receive_address = ""  '�ջ��˵�ַ
	receive_zip     = ""  '�ʱ�5 λ��6 λ�������
	receive_phone   = ""  '�ջ��˵绰
	receive_mobile  = ""  '�ջ����ֻ� ������11 λ����
	buyer_email     = ""  '��ҵ�֧�����˺�
    discount        = ""  '��Ʒ�ۿ�

 '�����Ҫ����Ӽ���������ʽ���������ӵڶ�����������,�������Ҫ������Ϊ��
   	logistics_fee_1	   = ""			'�������ͷ���  0.00
	logistics_payment_1  = ""	'�������ͷ��ø��ʽ��SELLER_PAY(����֧��)��BUYER_PAY(���֧��)��BUYER_PAY_AFTER_RECEIVE(��������)
	logistics_type_1	   = ""		'�������ͷ�ʽ��POST(ƽ��)��EMS(EMS)��EXPRESS(�������)


	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(service,subject,body,out_trade_no,price,quantity,seller_email,show_url,receive_name,receive_address,receive_zip,receive_phone,receive_mobile,buyer_email,discount,logistics_fee_1,logistics_payment_1,logistics_type_1)
	response.redirect itemUrl

	

	%>