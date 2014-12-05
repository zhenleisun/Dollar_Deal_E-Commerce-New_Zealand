

<table width="190" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><div align="center"><img src="images/body/left01.gif" width="190" height="43" /></div></td>
  </tr>
  <tr>
    <td style="background:url(images/leftbg.jpg) repeat-y;">
      <table  width="100%" border="0" cellpadding="0" cellspacing="0">
        <%set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 10 * from products where zhuangtai=0 order by chengjiaocount desc",conn,1,1
		%>
		<%i=0
		do while not rs.eof%>
        <tr>
          <td width="1%" height="24" valign="middle"></td>
          <td width="99%" valign="middle">&nbsp;<img src="images/body/dot_off.gif" width="23" height="13">&nbsp;&nbsp;
            <%response.write "<a href=products.asp?id="&rs("bookid")&"  title=successfully-sold-"&rs("chengjiaocount")&"-times >"
				if len(trim(rs("bookname")))>8 then
				response.write left(trim(rs("bookname")),16)&".."
				else
				response.write trim(rs("bookname"))
				end if
				response.write "</a>"
				%><font color="#DC143C">гд<%=trim(rs("huiyuanjia"))%></font>
            </td>
        </tr>
        <%i=i+1
		if i>=10 then exit do
		rs.movenext
		loop
		rs.close
		set rs=nothing%>
      </table></td>
  </tr>
</table>
