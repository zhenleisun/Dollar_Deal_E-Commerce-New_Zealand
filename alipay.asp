<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="alipayto/alipay_payto.asp"-->
 <%  qian=trim(request("p"))
 
 d=trim(request("d"))

	 shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '客户网站订单号，（现取系统时间，可改成网站自己的变量）
'必须的参数	
    service         =   "create_partner_trade_by_buyer"   'trade_create_by_buyer 表示标准双接口， create_partner_trade_by_buyer 表示担保交易接口
	subject			=	d '商品名称，
	body			=	webname&"订单"		'商品描述
	out_trade_no    =   dingdan        '按时间获取的订单号，可以修改成自己网站的订单号，保证每次提交的都是唯一即可
	

     price		= qian	'商品价格

    quantity        =   "1"             '商品数量,如果走购物车默认为1
    seller_email    =    seller_email   '卖家 的支付宝帐号，c2c客户，可以更改此参数。

 '以下是可选的参数 如果没有可以为空。注意：姓名、联系地址和邮政编码 这三项要么都为空，要么都不能为空。
  	show_url        = ""  '商户的展示地址，链接后面不能自定义参数
	receive_name    = ""  '收货人姓名
    receive_address = ""  '收货人地址
	receive_zip     = ""  '邮编5 位戒6 位数字组成
	receive_phone   = ""  '收货人电话
	receive_mobile  = ""  '收货人手机 必须是11 位数字
	buyer_email     = ""  '买家的支付宝账号
    discount        = ""  '商品折扣

 '如果需要多添加几组物流方式，可以增加第二组物流参数,如果不需要，可以为空
   	logistics_fee_1	   = ""			'物流配送费用  0.00
	logistics_payment_1  = ""	'物流配送费用付款方式：SELLER_PAY(卖家支付)、BUYER_PAY(买家支付)、BUYER_PAY_AFTER_RECEIVE(货到付款)
	logistics_type_1	   = ""		'物流配送方式：POST(平邮)、EMS(EMS)、EXPRESS(其他快递)


	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(service,subject,body,out_trade_no,price,quantity,seller_email,show_url,receive_name,receive_address,receive_zip,receive_phone,receive_mobile,buyer_email,discount,logistics_fee_1,logistics_payment_1,logistics_type_1)
	response.redirect itemUrl

	

	%>