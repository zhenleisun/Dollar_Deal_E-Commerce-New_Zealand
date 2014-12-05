<!--#include file="conn.asp"-->
<%
username=request.cookies("Cnhww")("username")
'shjianame=request.cookies("Cnhww")("shjianame")
response.Cookies("shangcheng").Expires =  NOW() -1
response.cookies("Cnhww")("username")=""
response.cookies("Cnhww")("shjianame")=""
response.cookies("Cnhww")("jifen")=0
response.cookies("Cnhww")("yucun")=0
response.cookies("Cnhww")("reglx")=0
response.cookies("Cnhww")("jiaoyijine")=0
response.Cookies("cnhww")("dingdanusername")=""

	session("userid")=""
conn.execute("delete from orders where username='"&username&"' and zhuangtai=7")
conn.execute("delete from ordersaward where username='"&username&"' and zhuangtai=7")
'response.redirect index.asp
response.Write "<script language=javascript>alert('You have successfully logged off£¡');window.location.href='index.asp';</script>"
%>
