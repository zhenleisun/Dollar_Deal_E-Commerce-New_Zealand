<!--#Include file="conn.asp"-->
<%
exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      checkid=rs("paytype")

response.redirect "send.asp"

%>