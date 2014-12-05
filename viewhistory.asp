<table width="190" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/viewhist.gif" width="190" height="50"></td>
  </tr>
  <tr>
    <td background="images/leftbg.jpg"><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" >
	  <%set rs=server.createobject("adodb.recordset")
		ProductList = request.cookies("leisure")("fav")
		If ProductList<>"" then
		sql="Select * From products Where bookid In (" & ProductList & ")"
		rs.open sql,conn,1,1
		if rs.eof and  rs.bof then 
		response.write "<div align=center>You are not currently browsing any goods</div>"
		end if 
		do while not rs.eof 
		%>
		<tr>
		  <td width="21%" align="center" ><img src="images/ring01.gif"></td> 
		  <td width="79%" height="30" align="left" ><a href="products.asp?id=<%=rs("bookid")%>"> 
	      <%response.write left(trim(rs("bookname")),20)%></a></td>
	  </tr>
	  <%rs.movenext
		loop
		else
		response.write "<tr><td align=center>You are not currently browsing any goods</td></tr>"
		end if 
		%>
	</table></td>
  </tr>
</table>







		