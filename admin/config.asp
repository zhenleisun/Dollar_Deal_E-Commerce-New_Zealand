<%
dim webname
set rs=server.CreateObject("adodb.recordset")
rs.Open "select webname from webinfo where webname='"&webname&"'",conn,1,1
webname=trim(rs("webname"))
rs.Close
set rs=nothing

%>
