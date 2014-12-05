<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<%
action=request("action")
id=request("id")

set rs=server.createobject("adodb.recordset")
if action="sj" then 
rs.open "select  * from products where bookid="&id ,conn,1,3
rs("zhuangtai")=0
rs.update
response.redirect request.servervariables("http_referer")
elseif action="xj" then

rs.open "select  * from products where bookid="&id ,conn,1,3
rs("zhuangtai")=1
rs.update
response.redirect request.servervariables("http_referer")
end if 
%>












