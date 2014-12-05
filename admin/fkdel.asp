<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<%
set rs=server.createobject("adodb.recordset")
conn.execute "delete from guestbook where id="&trim(request.querystring("id"))
response.redirect "viewreturn.asp"
%>