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
<%dim nclassid,anclassid
nclassid=int(request("nclassid"))
anclassid=int(request("anclassid"))
set rs=server.CreateObject("adodb.recordset")
rs.open "select nclassid,anclassid from ssort where nclassid="&nclassid ,conn,1,3
rs("anclassid")=anclassid
rs.Update
rs.Close
set rs=nothing
response.Write "<script language='javascript'>alert('转移成功！');history.go(-1);</script>"
%>