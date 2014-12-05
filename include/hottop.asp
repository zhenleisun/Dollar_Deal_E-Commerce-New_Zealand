<style type="text/css">
<!--
.style4 {color: #FF5C99}
-->
</style>
<%set rs=server.CreateObject("adodb.recordset")
	  rs.open "select top 10 * from products where zhuangtai=0 order by liulancount desc",conn,1,1%>

<table width="180" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><img src="images/body/left04.gif" width="180" height="50" /></div></td>
  </tr>
  <tr>
    <td><table   width="100%" border="0" cellpadding="0" cellspacing="0">
      <%i=0
do while not rs.eof%>
      <tr>
        <td width="1%" height="24" valign="middle"></td>
        <td width="99%" valign="middle">&nbsp;<img src="images/body/dian3.gif" width="11" height="11" />
			<%response.write "<a href=products.asp?id="&rs("bookid")&"  title=此商品已成功销售"&rs("chengjiaocount")&"次>"
				if len(trim(rs("bookname")))>8then
				response.write left(trim(rs("bookname")),8)&".."
				else
				response.write trim(rs("bookname"))
				end if
				response.write "</a>"
			%>
			<font color="#DC143C"><font color="#DC143C">￥<%=trim(rs("huiyuanjia"))%></font></font>
            <table width="100%" height="1"  border="0" cellpadding="0" cellspacing="0" background="image/lineDotGray.gif">
              <tr>
                <td height="1" background="images/blank.gif"></td>
              </tr>
          </table></td>
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
