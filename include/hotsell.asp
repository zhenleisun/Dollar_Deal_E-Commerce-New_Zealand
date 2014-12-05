<style type="text/css">
<!--
.style4 {color: #FF0000}
-->
</style>
<table width="660"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="images/body/news.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="330" height="50">&nbsp;</td>
        <td width="330"><a 
            href="fenleiview.asp?lx=news"><font 
            color="#ff3300">最新商品</font></a>&nbsp;<a 
            href="fenleiview.asp?lx=hot"><font 
            color="#ff3300">热卖商品</font></a>&nbsp;<a 
            href="fenleiview.asp?lx=tejia"><font 
            color="#ff3300">特价产品</font></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <%set rs=server.CreateObject("adodb.recordset")
		rs.open "select Top 12 * from products where bestbook=1 order by bookid desc",conn,1,1
		if rs.eof and rs.bof then
		response.write "<center><br><font color=red size=2>对不起，暂无新品！</font></font>"
		'response.End
		else
		%>
        <%
		if not rs.eof then
		i=1
		do while not rs.eof%>
        <td height="110" valign="middle"><table width="220" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="1"  background="images/lineDotGray.gif"></td>
                </tr>
                <tr>
                  <td><table width="100%" height="40" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="5"  ><a href="products.asp?id=<%=rs("bookid")%>" target="_blank"> </a></td>
                      </tr>
                      <tr>
                        <td  >&nbsp;&nbsp;<a href="products.asp?id=<%=rs("bookid")%>" target="_blank"><%=rs("bookname")%></a></td>
                      </tr>
                      <tr>
                        <td >&nbsp;&nbsp;<font color="#009999">会员价:<%=rs("huiyuanjia")%>元</font> </td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td><table width="220" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="123"><table width="96" height="96" border="0" align="center" >
                            <tbody>
                              <tr>
                                <td width="96" height="96" align="center" valign="middle" bgcolor="#ffffff"><div align="center">
								<%if rs("bookpic")="" then 
	response.write "<div align=center><a href=products.asp?id="&rs("bookid")&" ><img src=images/emptybook.gif width=75 height=75 border=0></a></div>"
	else%>
              <a href=products.asp?id=<%=rs("bookid")%> ><img src="<%=trim(rs("bookpic"))%>" width=96 height=96 border=0 align=absmiddle></a>
              <%end if%>
								</div></td>
                              </tr>
                            </tbody>
                        </table></td>
                        <td width="97" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="23%"><img src="images/Bullet_MoreInfo.gif" width="22" height="14" /></td>
                              <td width="77%">浏览</td>
                            </tr>
                            <tr>
                              <td colspan="2"><img src="images/btn_addtocart_small.gif" ></td>
                            </tr>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table></td>
              <td valign="middle" ><img src="images/body/line.gif" /></td>
            </tr>
          </table>
          <br /></td>
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
    </table></td>
  </tr>
</table>
