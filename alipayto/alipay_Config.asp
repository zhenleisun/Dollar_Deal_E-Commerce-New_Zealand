
<%
set rsb=server.CreateObject("adodb.recordset")
rsb.open "select * from webinfo ",conn,1,3


	seller_email	= rsb("alipayid")	 '请填写签约支付宝账号，
	partner			= rsb("partner")	 '填写签约支付宝账号对应的partnerID，
	key			    = rsb("alipaymd5")	 '填写签约账号对应的安全校验码

    notify_url		= "./Alipay_Notify.asp"	        '交易过程中服务器通知的页面 要用 http://格式的完整路径，例如http://www.alipay.com/alipay/Alipay_Notify.asp  注意文件位置请填写正确。
	return_url		= "./return_Alipay_Notify.asp"	'付完款后跳转的页面 要用 http://格式的完整路径, 例如http://www.alipay.com/alipay/return_Alipay_Notify.asp  注意文件位置请填写正确。
	'如果使用了Alipay_Notify.asp或者return_Alipay_Notify.asp，请在这两个文件中添加相应的合作身份者ID和安全校验码
	logistics_fee	   = "0.00"			'物流配送费用
	logistics_payment  = "BUYER_PAY"	'物流配送费用付款方式：SELLER_PAY(卖家支付)、BUYER_PAY(买家支付)、BUYER_PAY_AFTER_RECEIVE(货到付款)
	logistics_type	   = "EXPRESS"		'物流配送方式：POST(平邮)、EMS(EMS)、EXPRESS(其他快递)

	 	 
'登陆 www.alipay.com 后, 点商家服务,可以看到支付宝安全校验码和合作id,导航栏的下面 
%>