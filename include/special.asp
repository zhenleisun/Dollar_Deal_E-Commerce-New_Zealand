<TABLE width="961" border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR>
      <TD><img src="images/ds_18.jpg" width="961" height="81" border="0" usemap="#Map2" /></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr>
    <%set rs=server.CreateObject("adodb.recordset")
	rs.open "select Top "&tejia&" * from products where tejiabook=1 and zhuangtai=0 order by bookid desc ",conn,1,1
	if rs.eof and rs.bof then
	response.write "<center><br><font color=red size=2>Sorry, there are no such goods£¡</font></font>"
	'response.End
    else
	if not rs.eof then
	i=1
	do while not rs.eof
	%>
    <td valign="top" height="397" style="background:url(images/inprobg.jpg) no-repeat center center;"><table width="321" border="0" align="center" cellpadding="0" cellspacing="0">
			  <tr>
				<td height="58" align="left" valign="bottom"><div style=" padding:0 0 14px 25px; text-align:left;"><%response.write "<a href=products.asp?id="&rs("bookid")&">"
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
  end if
  end if 
  %>
</table>
<map name="Map2" id="Map2">
  <area shape="rect" coords="861,30,948,55" href="class.asp?lx=tejia" />
</map>