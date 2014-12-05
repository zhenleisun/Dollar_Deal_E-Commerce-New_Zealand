

<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/ds_15.jpg" width="961" height="76" /></td>
  </tr>
  <tr>
    <td width="321" align="center" valign="top">
      <table border="0" cellpadding="0" cellspacing="0">
        <%set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 6 * from products where zhuangtai=0 order by chengjiaocount desc, bookid asc",conn,1,1
		if rs.eof and rs.bof then
		response.write "<center><br><font color=red size=2>Sorry, there are no such goods！</font></center>"
		'response.End
		else
		%>
		<tr>
		<%
		if not rs.eof then
		i=1
		do while not rs.eof%>
          <td height="397" align="center" valign="top" style="background:url(images/inprobg.jpg) no-repeat center center;">
            <table width="321" border="0" align="center" cellpadding="0" cellspacing="0">
			  <tr>
				<td height="58" align="left" valign="bottom"><div style=" padding:0 0 14px 25px; text-align:left;"><%response.write "<a href=products.asp?id="&rs("bookid")&"  title=此商品已成功销售"&rs("chengjiaocount")&"次>"
				if len(trim(rs("bookname")))>12 then
				response.write left(trim(rs("bookname")),36)&".."
				else
				response.write trim(rs("bookname"))
				end if
				response.write "</a>"
				%></div></td>
			  </tr>
			  <tr>
				<td height="190" align="center" valign="top"><%if rs("bookpic")="" then 
				response.write "<div align=center><a href=list.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
				else%>
				<a href="Products.asp?id=<%=rs("bookid")%>" target="_blank"> <img src="<%=trim(rs("bookpic"))%>" width="278" height="180" border="0" align="absmiddle" /></a>
				<%end if%></td>
			  </tr>
			  <tr>
				<td height="56" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="10%" height="48">&nbsp;</td>
                    <td width="39%" align="left" valign="top" style="font-size:24px; color:#000;">$<%=trim(rs("huiyuanjia"))%></font></td>
                    <td width="51%" align="left">&nbsp;</td>
                  </tr>
                  
                </table></td>
			  </tr>
			  <tr>
				<td height="75" align="left" valign="middle"><table width="321" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="28" height="31">&nbsp;</td>
                    <td width="124" align="center" valign="top"><a style=" font-size:18px;" href="buy.asp?id=<%=rs("bookid")%>&amp;action=add">Buy</a></td>
                    <td width="12">&nbsp;</td>
                    <td width="126" align="center" valign="top"><a  style=" font-size:18px;" class="view" href="Products.asp?id=<%=rs("bookid")%>" target="_blank">View</a></td>
                    <td width="31">&nbsp;</td>
                  </tr>
                </table></td>
			  </tr>
			</table>
          </td>
		  <%if i mod 3 = 0 then%>
        </tr>
        <%end if
		  rs.movenext
		  i=i+1
		  loop
		rs.close
		set rs=nothing
		end if
  		end if %>
    </table></td>
  </tr>
</table>
