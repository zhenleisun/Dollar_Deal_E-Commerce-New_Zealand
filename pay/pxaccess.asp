<%@ Language=VBScript %>
<%

Dim sXmlAction,returnUrl,pxPayUrl,Userid,Key
returnUrl ="http://www.baidu.com"
pxPayUrl = "https://sec.paymentexpress.com/pxaccess/pxpay.aspx"
Userid = "DollarDeal_Dev"
Key = "6748d439ff1ae7afc5b0eb02bbf9c10e32337cdb6919254811feee4db35b236c"

If Request.Form("AmountInput") <> "" Then

sXmlAction = sXmlAction & "<GenerateRequest><PxPayUserId>" & Userid & "</PxPayUserId>"
sXmlAction = sXmlAction & "<PxPayKey>" & Key & "</PxPayKey>"
sXmlAction = sXmlAction & "<TxnType>" & Request.Form("TxnType") & "</TxnType>"
sXmlAction = sXmlAction & "<CurrencyInput>NZD</CurrencyInput>"
sXmlAction = sXmlAction & "<MerchantReference>" & Request.Form("MerchantRef") & "</MerchantReference>"
sXmlAction = sXmlAction & "<AmountInput>" & Request.Form("AmountInput") & "</AmountInput>"
sXmlAction = sXmlAction & "<UrlFail>" & returnUrl & "</UrlFail>"
sXmlAction = sXmlAction & "<UrlSuccess>" & returnUrl & "</UrlSuccess>"
sXmlAction = sXmlAction & "</GenerateRequest>"

Dim objXMLhttp
Set objXMLhttp = server.Createobject("MSXML2.XMLHTTP")

objXMLhttp.Open "POST", pxPayUrl ,False
objXMLhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objXMLhttp.send sXmlAction
response.write "start xml dom"
Set xml = Server.CreateObject("Microsoft.XMLDOM")
response.write "done xml dom"
xml.async = False
xml.LoadXml (objXMLhttp.responsetext)

response.Redirect xml.documentElement.childNodes(0).text

Set objXMLhttp = nothing

End If
If Request.QueryString("Result") <> "" Then

sXmlAction = sXmlAction & "<ProcessResponse><PxPayUserId>"& Userid &"</PxPayUserId>"
sXmlAction = sXmlAction & "<PxPayKey>"& Key &"</PxPayKey>"
sXmlAction = sXmlAction & "<Response>" & Request.QueryString("Result") & "</Response>"
sXmlAction = sXmlAction & "</ProcessResponse>"

Dim objXML
Set objXML = server.Createobject("MSXML2.XMLHTTP")
objXML.Open "POST", pxPayUrl ,False
objXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
objXML.send sXmlAction
response.write Server.HTMLEncode(objXML.responsetext)

Set objXML = nothing

End If
%>