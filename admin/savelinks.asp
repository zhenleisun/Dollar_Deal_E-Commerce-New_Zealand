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
 dim action,linkid
linkid=request.QueryString("id")
if linkid<>"" then
if not isnumeric(linkid) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from links",conn,1,3
rs.AddNew
rs("linkname")=trim(request("linkname1"))
rs("linkurl")=trim(request("linkurl1"))
rs("linkidorder")=int(request("linkidorder1"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "links.asp"
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from links where linkid="&linkid,conn,1,3
rs("linkname")=request("linkname")
rs("linkurl")=trim(request("linkurl"))
rs("linkidorder")=int(request("linkidorder"))
rs.update
rs.close
response.Redirect "links.asp"
set rs=nothing
case "del"
conn.execute "delete from links where linkid="&linkid
response.Redirect "links.asp"
end select
%>
