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
<%dim action,url,i,abc,anclassid,anclass
anclassid=request("anclassid")
anclass=request.QueryString("anclass")
url="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ssort",conn,1,3
rs.AddNew
rs("nclass")=trim(request("nclass2"))
rs("nclassidorder")=int(request("nclassidorder2"))
rs("anclassid")=int(request("anclassid"))
rs.Update
rs.Close
set rs=nothing
response.redirect url&"managessort.asp?id="&anclassid&"&anclass="&anclass
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ssort where nclassid="&request.QueryString("id"),conn,1,3
rs("nclass")=trim(request("nclass"))
rs("nclassidorder")=int(request("nclassidorder"))
rs.update
rs.close
set rs=nothing
response.redirect url&"managessort.asp?id="&anclassid&"&anclass="&anclass
case "del"
anclassid=request.QueryString("anclassid")
conn.execute ("delete from ssort where nclassid="&request.QueryString("id"))
conn.execute ("delete from products where nclassid="&request.QueryString("id"))
response.redirect url&"managessort.asp?id="&anclassid&"&anclass="&anclass
end select
%>
<%
Function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>
