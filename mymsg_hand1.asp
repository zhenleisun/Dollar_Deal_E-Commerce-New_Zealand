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
	response.write "alert('�����ɹ����������Ա����һ��վ����Ϣ��');"
	response.write "location.href='user.asp?action=famess';"
	response.write "</script>"
	response.end

conn.close
set conn=nothing
%>

<p><span lang="zh-cn">���ͳɹ�</span></p>
