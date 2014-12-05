<table width="725"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr valign="middle"> 
  	<td height="10"></td>
  </tr>
  <tr>
    <td align="left"><table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" >
      <tr>
        <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select Top "&xp&" * from products where newsbook=1 and zhuangtai=0 order by bookid desc ",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "<center><br><font color=red size=2>Sorry, there are no such goods£¡</font></font>"
		  'response.End
		  else
		  %>
        <%
		if not rs.eof then
		i=1
		do while not rs.eof%>
        <td align="center" valign="top" ><table width="160" border="0" align="center" cellpadding="0" cellspacing="0" >
              <tr> 
                <td align="center"><table width="160" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tbody>
                    <tr>
                      <td height="100" bgcolor="#ffffff" align="center"><%if rs("bookpic")="" then 
	response.write "<div align=center><a href=list.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
	else%>
                          <a href="Products.asp?id=<%=rs("bookid")%>" target="_blank"> <img src="<%=trim(rs("bookpic"))%>" width="160" height="154" border="0" align="absmiddle" /></a>
                          <%end if%>
                      </td>
                    </tr>
                  </tbody>
                </table></td>
              </tr>
              <tr> 
                <td align="center"><table width="95%" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td align="left" >
                        <%response.write "<a class=a4 href=products.asp?id="&rs("bookid")&" ><font color=#666666>"
						if len(trim(rs("bookname")))>7 then
						response.write left(trim(rs("bookname")),25)&".."
						else
						response.write trim(rs("bookname"))
						end if
						response.write "</font></a>"
						%></td>
                    </tr>
                    <tr> 
                      <td align="left" style="font-size:14px; font-weight:bold; color:#000000">$ <%=rs("huiyuanjia")%></td>
                    </tr>
                  </table></td>
              </tr>
          </table></td>
        <%if i mod 4 = 0 then%>
      </tr>
      <%end if
	    rs.movenext
         i=i+1
    	 loop
		rs.close
		end if
		end if 
	%>
    </table></td>
  </tr>
</table>


<map name="Map4" id="Map4">
  <area shape="rect" coords="68,7,80,15" href="class.asp?lx=news" />
</map>