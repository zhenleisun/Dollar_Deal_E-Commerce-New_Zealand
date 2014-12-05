<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--我的收藏</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="5" marginwidth="0" marginheight="0">
<%
if request.cookies("Cnhww")("username")="" then
response.write "<script language=javascript>alert('Sorry, you are not logged in！');window.close();</script>"
response.End
end if
dim bookid,username,action
action=request.QueryString("action")
username=trim(request.cookies("Cnhww")("username"))
bookid=request.QueryString("id")
if InStr(action,"'")>0 then
response.write"<script>alert(""Illegal access!"");window.close();</script>"
response.end
end if
if bookid<>"" then
if not isnumeric(bookid) then 
response.write"<script>alert(""Illegal access!"");window.close();</script>"
response.end
else
if not isinteger(bookid) then
response.write"<script>alert(""Illegal access!"");window.close();</script>"
end if
end if
end if
select case action
case "del"
conn.execute "delete from orders where actionid="&request.QueryString("actionid")
if request.QueryString("ll")=1 then
response.redirect "user.asp?action=shoucang"
else
response.redirect "stow.asp?action=show"
end if
response.End
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username from orders where username='"&username&"' and bookid="&bookid&" and zhuangtai=6",conn,1,1
if not rs.eof and not rs.bof then
response.write "<script language=javascript>alert('Sorry, this product already exists in the collection frame you, can not be repeated to add！');window.location.href='stow.asp?action=show';</script>"
response.end
rs.close
set rs=nothing
else
if rs.recordcount=10 then
response.write "<script language=javascript>alert('Sorry, you can only have a maximum of 10 pieces of goods！');window.location.href='stow.asp?action=show';</script>"
response.end
else
rs.close
set rs=nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username,zhuangtai,zonger from orders",conn,1,3
rs.addnew
rs("bookid")=bookid
rs("username")=username
rs("zhuangtai")=6
rs("zonger")=0
rs.update
end if
rs.close
response.Redirect "stow.asp?action=show"
set rs=nothing
end if
case "show"
response.write "<table width=96% border=0 align=center cellpadding=2 cellspacing=2><tr><td width=60% >"
response.write "</td><td width=40% valign=baseline><div align=right>Sorry, you can only have a maximum of 10 pieces of goods</div></td></tr></table>"
set rs=server.CreateObject("adodb.recordset")
rs.open "select orders.actionid,orders.bookid,products.bookname,products.shichangjia,products.huiyuanjia,products.dazhe from products inner join  orders on products.bookid=orders.bookid where orders.username='"&request.cookies("Cnhww")("username")&"' and orders.zhuangtai=6",conn,1,1 
%>
<table width=96% border=0 align=center cellpadding=2 cellspacing=1 bgcolor=#cccccc>
<form name='form1' method='post' action="stowcart.asp">
<%
response.write "<tr align=center><td width=8% bgcolor=#f1f1f1>select</td>"
response.Write "<td width=42% bgcolor=#f1f1f1>Name</td>"
response.Write "<td width=14% bgcolor=#f1f1f1>market price</td>"
response.Write "<td width=14% bgcolor=#f1f1f1>Member price</td>"
response.Write "<td width=8% bgcolor=#f1f1f1>Delete</td></tr>"
do while not rs.eof
response.write "<tr><td bgcolor=#ffffff><div align=center><input name=bookid type=checkbox checked value="&rs("bookid")&" ></div></td>"
response.write "<td bgcolor=#ffffff STYLE='PADDING-LEFT: 5px'><div align=left><a href=products.aspid="&rs("bookid")&" >"&rs("bookname")&"</a></div></td>"		  
response.write "<td bgcolor=#ffffff><div align=center>$ "&formatnumber(rs("shichangjia"),0)&"</div></td>"	
response.write "<td bgcolor=#ffffff><div align=center><font color=#dd6600>$ "&formatnumber(rs("huiyuanjia"),0)&"</font></div></td>"
response.write "<td bgcolor=#ffffff><div align=center>"
response.Write "<a href=stow.asp?action=del&actionid="&rs("actionid")&">"
response.Write "<img src=images/trash.gif border=0></a></div></td></tr>"
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr>
<td height=36 colspan=6 bgcolor=#ffffff align=center> 
<input class="go-wenbenkuang" onFocus="this.blur()" type="submit" name="submit" value=" The cashier ">
<input class="go-wenbenkuang" onFocus="this.blur()" onClick="javascript:window.close();" type=reset name="button" value=" Continue shopping ">
</td></tr></form></table>
<%end select%>
</body>
</html>
