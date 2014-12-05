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
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//添加新数据
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from bsort",conn,1,3
rs.AddNew
rs("anclass")=trim(request("anclass2"))
rs("anclassidorder")=int(request("anclassidorder2"))
rs("fudongjia")=int(request("fudongjia2"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "managebsort.asp"
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from bsort where anclassid="&anclassid,conn,1,3
rs("anclass")=trim(request("anclass"))
rs("anclassidorder")=int(request("anclassidorder"))
rs("fudongjia")=int(request("fudongjia"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "managebsort.asp"
case "del"
conn.execute ("delete from bsort where anclassid="&anclassid)
conn.execute ("delete from ssort where anclassid="&anclassid)
conn.execute ("delete from products where anclassid="&anclassid)
response.Redirect "managebsort.asp"
end select
%>
