<!--#include file="conn.asp"-->
<%

set rs=server.CreateObject("adodb.recordset")
rs.Open "select payid from webinfo ",conn,1,3
MerchantID_User=rs("westid")'Ö§¸¶ID(¿Í»§ID)
rs.close
set rs=nothing
''''''''''''''''''''''''''''''''''''''''''
' ¸ÃÊ¾·¶³ÌÐò½ÓÊÕÎ÷²¿Ö§¸¶·¢À´µÄÖ§¸¶³É¹¦Í¨Öª£¬ÑéÖ¤ÆäÕæÊµÐÔ£¬²¢×÷ÏàÓ¦µÄ×Ô¶¯´¦Àí
''''''''''''''''''''''''''''''''''''''''''

dim MerchantOrderNumber, WestPayOrderNumber, PaidAmount, MerchantID

MerchantOrderNumber=request("MerchantOrderNumber") 'ºÍÉÌ»§Ö§¸¶ÃüÁîÖÐµÄ¶©µ¥ºÅÏàÍ¬
WestPayOrderNumber=request("WestPayOrderNumber")
PaidAmount=ccur("0"&request("PaidAmount"))  'WestPay´«»ØµÄÊµ¼ÊÖ§¸¶½ð¶î¡£ÓÃCCUR×ªÎªÊý×ÖÐÍ¡£
'×¢£ºÉÌ»§±ØÐë¸ù¾ÝÎÒÃÇ´«»ØÉÌ»§Ô­Ê¼¶©µ¥ºÅÕÒµ½Ô­Ê¼¶©µ¥£¬±È½ÏÊµ¸¶½ð¶îºÍÔ­Ê¼¶©µ¥½ð¶î£¬ÏàÍ¬²ÅÊÇÖ§¸¶³É¹¦¡£
MerchantID=request("MerchantID")
'×¢£ºÉÌ»§±ØÐëÅÐ¶Ï´ËÉÌ»§IDÊÇ²»ÊÇÄúµÄÉÌ»§ID

Dim objHttp, str

Dim PaymentCompleted
PaymentCompleted=false

' ×¼±¸»Ø´«Ö§¸¶Í¨Öª±íµ¥
str = Request.Form & "&cmd=validate"

'set objHttp = Server.CreateObject("Msxml21.ServerXMLHTTP")
'set objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")
'set objHttp = Server.CreateObject("Microsoft.XMLHTTP")
set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
 
'°ÑWestPay´«À´µÄÍ¨ÖªÄÚÈÝÔÙ´«»Øµ½WestPay×÷ÑéÖ¤ÒÔÈ·±£Í¨ÖªÐÅÏ¢µÄÕæÊµÐÔ 
 
objHttp.open "POST", "http://www.WestPay.com.cn/pay/ISPN.asp", false
'ISPN: Instant Secure Payment Notification

objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHttp.Send str

if (objHttp.status <> 200 ) then
	    'HTTP ´íÎó´¦Àí
		response.Write("Status="&objHttp.status)
elseif (objHttp.responseText = "VERIFIED") then
	
		'Ö§¸¶Í¨ÖªÑéÖ¤³É¹¦
		if MerchantID=""&MerchantID_User&"" then 'ÅÐ¶Ï´Ë¶©µ¥ÊÇ²»ÊÇÉÌ»§µÄ¶©µ¥¡£
			PaymentCompleted=true
		end if
	
elseif (objHttp.responseText = "INVALID") then
		'Ö§¸¶Í¨ÖªÑéÖ¤Ê§°Ü
		response.Write("Invalid")
else
		'Ö§¸¶Í¨ÖªÑéÖ¤¹ý³ÌÖÐ³öÏÖ´íÎó
		response.Write("Error")
end if
set objHttp = nothing

if PaymentCompleted then
	'Ö§¸¶³É¹¦µÄ´¦Àí...£¨ÒÔÏÂ³ÌÐò½ö¹©²Î¿¼£©

	'×¢£ºÉÌ»§±ØÐë¸ù¾ÝÎÒÃÇ´«»ØÉÌ»§Ô­Ê¼¶©µ¥ºÅÕÒµ½Ô­Ê¼¶©µ¥£¬±È½ÏÊµ¸¶½ð¶îºÍÔ­Ê¼¶©µ¥½ð¶î£¬ÏàÍ¬²ÅÊÇÖ§¸¶³É¹¦¡£
	'´Ë´¦ÎªÔÚÏß³äÖµ,²»ÓÃ½ð¶î±È¶ÔÊÇ·ñ¹»!
	
	'if ÄúµÄÔ­Ê¼¶©µ¥½ð¶î=PaidAmount then
	reseponse.write "¹§Ï²£¬Ö§¸¶³É¹¦!"
	'else
		'Ê§°Ü
	'end if
else

	'Ö§¸¶²»³É¹¦µÄ´¦Àí...
response.redirect "logout3.asp"end if
%>ÿ