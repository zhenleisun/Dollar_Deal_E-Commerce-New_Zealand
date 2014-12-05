<%set rs=server.CreateObject("adodb.recordset")
	  rs.open "select top 10 * from products where pingjizong>0 order by pingji desc,adddate",conn,1,1%>
<style type="text/css">
<!--
.style4 {color: #FF5C99}
-->
</style>
<table width="160" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/body/reviwtop.gif" width="160" height="37" /></td>
  </tr>
  <tr>
    <td><table style=" BORDER-bottom:#10ACC6 1px solid;BORDER-left:#10ACC6 1px solid; BORDER-right:#10ACC6 1px solid;" width="100%" border="0" cellpadding="0" cellspacing="0">
      <%i=0
do while not rs.eof%>
      <tr>
        <td width="1%" height="24" valign="middle"></td>
        <td width="99%" valign="middle">&nbsp;<img src="images/body/dian4.gif" width="15" height="9" />
            <%response.write "<a href=products.asp?id="&rs("bookid")&"  title=此商品已被浏览了"&rs("liulancount")&"次>"
				if len(trim(rs("bookname")))>9 then
				response.write left(trim(rs("bookname")),7)&".."
				else
				response.write trim(rs("bookname"))
				end if
				response.write "</a>"
				%>
            <font color="#DC143C">↑<font color="#DC143C">￥<%=trim(rs("huiyuanjia"))%></font></font>
            <table  width="100%" height="1"  border="0" cellpadding="0" cellspacing="0" background="image/lineDotGray.gif">
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
