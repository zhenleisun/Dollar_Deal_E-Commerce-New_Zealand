<!--#include file="conn.asp"-->
<%
if request.form("name")="" then
response.Write "<script LANGUAGE='javascript'>alert('Please fill in your name');history.go(-1);</script>"
response.End
end if
if request.form("content")="" then
response.Write "<script LANGUAGE='javascript'>alert('Please fill in your message and suggestions');history.go(-1);</script>"
response.End
end if

if request.form("temp")="" then
response.write "<script language='javascript'>" & VbCRlf
response.write "alert('Illegal operation!');" & VbCrlf
response.write "history.go(-1);" & vbCrlf
response.write "</script>" & VbCRLF
'由于表单被重复提交(标志为session("antry")为空)引起
else
iname=request.Form("name")
iqq=request.Form("qq")
iemail=request.Form("email")
iurl=request.form("url")
isex=request.form("sex")
icontent=request.form("content")
%>

<%
set rss=Server.CreateObject("ADODB.RecordSet")
	sqls="select * from webinfo"
	rss.open sqls,conn,1,3
        view=rss("view")
%>
<%
set rs=server.createobject("adodb.recordset")
rs.open "select * from guestbook where (id is null)",conn,1,3
rs.addnew
rs("name")=iname
rs("qq")=iqq
rs("email")=iemail
rs("url")=iurl
rs("sex")=isex
rs("content")=icontent
rs("time")=now()
if view<>"0" then view="1"
			rs("online")=view
rs.update
rs.close
set rs=nothing
session("antry")="" '提交成功，清空session("antry")，以防重复提交！！
end if
%>
<meta http-equiv="refresh" content="1;URL=viewreturn.asp">
<title>Message success</title>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="62%" border="0" align="center">
  <tr>
    <td align="center"><p> <img src="images/loading.gif" width="94" height="27"><br>
        <br><%
  	if view="0" then
	response.write "The message submitted successfully, the message must be approved by the administrator to issue audit."
	else
	response.write "The message submitted successfully, click 'OK' to return to the message list！"
	end if
	%>
      </td>
  </tr>
</table>