<!--#include file="conn.asp"-->
<%

action=request.QueryString("action")
if request.Cookies("cnhww")("username")<>"" then
username=trim(request.Cookies("cnhww")("username"))
else
if request.Cookies("cnhww")("dingdanusername")="" then
username=now()
username=replace(trim(username),"-","")
username=replace(username,":","")
username=replace(username," ","")
response.Cookies("cnhww")("dingdanusername")=username
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] ",conn,1,3
rs.addnew
rs("username")=username
rs("niming")=1
rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
rs("adddate")=now()
rs("lastlogin")=now()
rs("logins")=1
rs.update
rs.close
set rs=nothing

else
username=request.Cookies("cnhww")("dingdanusername")
end if
end if
bookid=request.QueryString("id")







if InStr(action,"'")>0 then
response.write"<script>alert(""Illegal access!"");window.close();</script>"
response.end
end if
if bookid<>"" then
if not isnumeric(bookid) then 
response.write"<script>alert(""Illegal access!"");window.close();</script>"
response.end
else
if not isinteger(bookid) then
response.write"<script>alert(""Illegal access!"");window.close();</script>"
end if
end if
end if
select case action
case "del"
conn.execute "delete from orders where actionid="&request.QueryString("actionid")
if request.QueryString("ll")=22 then
response.redirect "myuser.asp?action=shoucang"
else
response.redirect "buy.asp?action=show"
end if
response.End
case "add"
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from products where bookid="&bookid,conn,1,1
if request.Cookies("cnhww")("reglx")=2 then 
	danjia=rs_s("vipjia")
else
	danjia=rs_s("huiyuanjia")
end if
kucun=rs_s("kucun")
bookname=rs_s("bookname")
shjiaid=rs_s("shjiaid")
rs_s.close
set rs_s=nothing
if kucun<=0 then
response.write "<script language=javascript>alert('You buy the product " &bookname& " out of stock for the time being not in the shopping cart, please buy other goods£¡');location.href='javascript:onclick=history.go(-1)'</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username,bookcount,zonger from orders where username='"&username&"' and bookid="&bookid&" and zhuangtai=7",conn,1,3
if rs.recordcount=1 then
if kucun<(rs("bookcount")+1) then
response.write "<script language=javascript>alert('You buy the product "&bookname&" out of stock for the time being not in the shopping cart, please buy other goods£¡');location.href='javascript:onclick=history.go(-1)'</script>"
response.end
end if
rs("zonger")=(rs("bookcount")+1)*danjia
rs("bookcount")=rs("bookcount")+1
rs.update
rs.close
set rs=nothing
response.Redirect "buy.asp?action=show"
else
rs.close
set rs=server.CreateObject("adodb.recordset")
rs.open "select bookid,username,shjiaid,zhuangtai,zonger,bookcount,niming,chima,yanse from orders",conn,1,3
rs.addnew
rs("bookid")=bookid
rs("username")=username
rs("zhuangtai")=7
rs("bookcount")=1
rs("chima")=request.form("chima")
rs("yanse")=request.form("yanse")
rs("shjiaid")=shjiaid
rs("zonger")=danjia
if request.cookies("Cnhww")("username")="" then
rs("niming")=1
end if
rs.update
rs.close
set rs=nothing
response.Redirect "buy.asp?action=show"
end if
case "show"
%>
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--My cart</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/ds_07.jpg); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="userinfo.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td><!--#include file="shopcart.asp"--></td>
      </tr>
      <tr>
        <td><!--#include file="include/selltop.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><TABLE cellSpacing=0 cellPadding=0 width=100% align=center border=0>
      <TBODY>
        <TR><%
		set rs=server.CreateObject("adodb.recordset")
rs.open "select orders.actionid,orders.bookid,orders.bookcount,orders.zonger,orders.shjiaid,products.bookname,products.shichangjia,products.huiyuanjia,products.vipjia,orders.chima,orders.yanse from products inner join  orders on products.bookid=orders.bookid where orders.username='"&username&"' and orders.zhuangtai=7",conn,1,1 
shuliang=rs.recordcount
%>
            <TD class=b vAlign=top> 
              <%
response.write "<table width=100% align=center border=0 cellspacing=0 cellpadding=0  background=images/body/pdbg01.gif><tr><td  height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp>"&webname&"</a> >> My Cart</td</tr></table>"
 
%>
              <% if shuliang=0 then %><br><div align=center>Your shopping cart is empty, please thanks for shopping!</div><%end if %></div></div></div>
            <br>
            <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#f1f1f1" >
                <form name='form1' method='post' action=chcount.asp>
                  <tr bgcolor=#f1f1f1 align=center> 
                    <td width=25%  >Product Name</td>
                    <td width=16% >Price 
                      <%if request.Cookies("cnhww")("reglx")=2 then %>
                      (VIP) 
                      <%else%>
                      (Member) 
                    <%end if%>                    </td>
                    <td width=12% >Number</td>
                    <td width="9%">Size</td>
                    <td width="13%">Color</td>
                    <td width=15% >The total price</td>
                    <td width=10% >Delete</td>
                  </tr>
                  <%
jianshu=0
zongji=0
do while not rs.eof%>
                  <tr bgcolor="#FFFFFF"> 
                    <td width="25%" height="30" align="center" ><a href=products.asp?id=<%=rs("bookid")%> ><%=rs("bookname")%></a> 
                      <input name=bookid type=hidden value="<%=rs("bookid")%>"> 
                      <input name=actionid type=hidden value="<%=rs("actionid")%>"></td>
                    <td width="16%" height="30" align="center" > 
                      $
                        <%
	if request.Cookies("cnhww")("reglx")=2 then 
	response.write rs("vipjia")
	else
	response.write rs("huiyuanjia")
	end if%>
                    </td>
                    <td width="12%" height="30" align="center" ><input class=wenbenkuang name=bookcount value="<%=rs("bookcount")%>" size=5>                    </td>
                    <td height="30" align="center" ><%=rs("chima")%>&nbsp;</td>
                    <td height="30" align="center" ><%=rs("yanse")%></td>
                    <td align="center" width="15%" height="30" >$<font color=red><%=rs("zonger")%></font> 
                    </td>
                    <td align="center" width="10%" height="30" ><a href=buy.asp?action=del&actionid=<%=rs("actionid")%>><img src=images/trash.gif border=0></a>                    </td>
                  </tr>
                  <%
jianshu=jianshu+rs("bookcount")
zongji=zongji+rs("zonger")
rs.movenext
loop
rs.close
set rs=nothing%>
                  <tr bgcolor="#FFFFFF"> 
                    <td height=30 colspan=7 align=center >There are <font color=red><%=shuliang%></font> kinds of goods in the shopping cart&nbsp;&nbsp;A total of <font color=red><%=jianshu%></font>&nbsp;parts&nbsp;&nbsp;Total $<font color=red><%=zongji%></font>&nbsp;&nbsp;&nbsp;You have a deposit $£º<%=request.cookies("Cnhww")("yucun")%>  </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td align="center" height=50 colspan=7><input class="go-wenbenkuang" type="button" name="imageField12" value="Continue shopping"  border="0" onFocus="this.blur()" onClick="this.form.action='index.asp';this.form.submit()"> 
                      <input class="go-wenbenkuang" type="button" name="imageField2" value="Change number"  border="0" onFocus="this.blur()" onClick="this.form.action='chcount.asp';this.form.submit()"> 
                      <input name="imageField22" class="go-wenbenkuang" type="button" value="Empty the shopping cart" onFocus="this.blur()" onClick="this.form.action='clearcart.asp';this.form.submit()"> 
                      <input name="imageField222" class="go-wenbenkuang" type="button" value="The cashier" onFocus="this.blur()" onClick="this.form.action='mycart.asp';this.form.submit()"></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td height="60" colspan="7" STYLE="PADDING-LEFT: 20px">
					  ¡¤ If you would like to continue shopping, please click to continue shopping<br>
                      ¡¤ If you want to update in a shopping cart product, please change, and then select a revision number<br>
                      ¡¤ If you want to cancel all the ordered products in the shopping cart, please click to empty the shopping cart<br>
                      ¡¤ If you are satisfied with your purchase of the product, please check to the cashier (member must first login, non member must first free registration become the member)<br> </td>
                  </tr>
                </form>
              </table></TD>
        </TR>
      </TBODY>
    </TABLE></td>
 </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
<%end select%>
