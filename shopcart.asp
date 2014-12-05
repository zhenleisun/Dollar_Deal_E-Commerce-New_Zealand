<table width="190" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td ><img src="images/body/left05.gif" width="190" height="38"  /></td>
</tr>
<tr>
  <td style="background:url(images/leftbg.jpg) repeat-y;"><table width="100%" border="0" cellpadding="0" cellspacing="0" align="center">
    <%
	if request.cookies("Cnhww")("username")<>"" then 
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select count(*) as rec_count from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai=7",conn,1,1
	rec_count=rs("rec_count")
	rs.close
	rs.open "select sum(zonger) as zongji from orders where username='"&request.cookies("Cnhww")("username")&"' and zhuangtai=7",conn,1,1
	%>
    <tr>
      <td height="2" align="cneter"></td>
    </tr>
    <tr>
      <td height="20">　Cart is <font color="red"><%=rec_count%></font> pieces of goods</td>
    </tr>
    <tr>
      <td height="20">　sum of money��$<font color="red"><%=rs("zongji")%></font></td>
    </tr>
    <tr>
      <td height="20">　<a href="buy.asp?action=show&amp;lx=1" target="_blank">View the cart/checkout&gt;&gt;</a></td>
    </tr>
    <tr>
      <td height="20">　<a href="user.asp?action=shoucang">My collection&gt;&gt;</a></td>
    </tr>
    <%rs.close
		set rs=nothing
		else%>
    <tr>
      <td height="20"><div align="center">You had<font color="red"> not Login</font><br />
        shopping cart <font color="red">can not use</font> </div></td>
    </tr>
    <tr>
      <td height="1"></td>
    </tr>
    <%end if%>
  </table></td>
</tr>
</table>
