<!--#include file="conn.asp"-->
<%
if request.cookies("Cnhww")("username")<>"" then
username=trim(request.cookies("Cnhww")("username"))
else
username=request.cookies("Cnhww")("dingdanusername")
end if
conn.execute("delete from orders where username='"&username&"' and zhuangtai=7")
response.Redirect "buy.asp?action=show"
%>
