<TABLE width="741" 
border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR>
      <TD><img src="images/tjj.gif" width="741"   border="0" /></TD>
    </TR>
  </TBODY>
</TABLE>
<table width=741 border=0 align="center" cellpadding=0 cellspacing=0>
  <tr><td>                                                                                                                                                                                         
	 <div id=demo style="overflow:hidden;height:150;width:741;color:#ffffff"><table align=left cellpadding=0 cellspace=0 border=0><tr><td id=demo1 valign=top>
		<table border=0 cellpadding=0 cellspacing=0>
		<tr>
                  <td> 
                    <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select    * from products where tejiabook=1 and zhuangtai=0 order by bookid desc",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "<center><br><font color=red size=2>对不起，暂无精品！</font></font>"
		  'response.End
		  else
		  do while not rs.eof 
		  %>
<td height="110" valign="top" ><table width="128" align="center" cellpadding="0" cellspacing="0" >
      <tr valign="middle">
        <td></td>
      </tr>
      <tr>
        <td height="113" align="center"><table  width="140" height="100" cellspacing="1" cellpadding="2"  border="0">
              <tbody>
                <TR> 
                                <TD align=middle <% if probg=0 then %>
                      background="images/136.jpg" <%end if %>
                        height=140> 
                                  <%if rs("bookpic")="" then 
	response.write "<div align=center><a href=products.asp?id="&rs("bookid")&" ><img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
	else%>
                                  <a href="products.asp?id=<%=rs("bookid")%>" ><img src="<%=trim(rs("bookpic"))%>" alt="点击查看商品:<%=rs("bookname")%>" width=105 height=105 border="0" align="absmiddle" /></a> 
                                  <%end if%>
                                </TD>
                </TR>
              </tbody>
            </table></td>
      </tr>
      <tr>
        <td align="center"></td>
      </tr>
    </table>
    </td>
     <% rs.movenext
loop
end if 
%>
                </tr>
		</table>
	</td><td id=demo2 valign=top></td></tr></table></div>
  <script>
  var speed=10//速度数值越大速度越慢
  demo2.innerHTML=demo1.innerHTML
  function Marquee(){
  if(demo2.offsetWidth-demo.scrollLeft<=0)
  demo.scrollLeft-=demo1.offsetWidth
  else{
  demo.scrollLeft++
  }
  }
  var MyMar=setInterval(Marquee,speed)
  demo.onmouseover=function() {clearInterval(MyMar)}
  demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)}
  </script>
	</td></tr>                                                                                                                                                                                                                            
	</table>
