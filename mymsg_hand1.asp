<!--#include file="conn.asp"--><head>
<meta http-equiv="Content-Language" content="en-us">
</head>

<%


	sql="select * from sms"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
  	rs("name")="admin"
	neirong=request.form("neirong")
	neirong=replace(neirong,"=======","<font color=darkgray>======")
  	rs("neirong")=neirong
  	rs("riqi")=now()
 	rs("fname")=request.cookies("cnhww")("username")
	rs.update	
	rs.close
	set rs=nothing

	response.write "<script language='javascript'>"
	response.write "alert('操作成功，已向管理员发送一条站内消息！');"
	response.write "location.href='user.asp?action=famess';"
	response.write "</script>"
	response.end

conn.close
set conn=nothing
%>

<p><span lang="zh-cn">发送成功</span></p>
