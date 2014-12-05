<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head>
<%
if request.Cookies("cnhww")("username")<>"" then
username=trim(request.Cookies("cnhww")("username"))
else
if request.Cookies("cnhww")("dingdanusername")="" then
username=now()
username=replace(trim(username),"-","")
username=replace(username,":","")
username=replace(username," ","")
response.Cookies("cnhww")("dingdanusername")=username
username=request.Cookies("cnhww")("dingdanusername")

else
username=request.Cookies("cnhww")("dingdanusername")
end if 
end if
							
%>
<title><%=webname%>-商品对比</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="<%=des%>">
<meta name="keywords" content="<%=keya%>">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  style="background:url(images/ds_07.jpg); text-align:center;">
<%
dim shopid,action

action		= request.QueryString("action")
shopid		= request.QueryString("id")
actionid	= request.QueryString("actionid")
'//删除收藏
select case action
case "del"


conn.execute "delete from shop_action where actionid="&actionid
response.redirect "to.asp?action=show"
response.End
case "add"
'//对比，判断是否存在


set rs3=server.CreateObject("adodb.recordset")
rs3.open "select shopid from shop_action where shop_action.username='"&username&"'and shopid="& shopid &" and zhuangtai=8",conn,1,1
if not rs3.eof and not rs3.bof then
response.write "<script language=javascript>alert('对不起，此商品已存在于您的对比架中，不可以重复添加！');window.location.href='to.asp?action=show';</script>"
response.end
end if
rs3.close
set rs3=Nothing

'//判断对比数
 set rs3=server.CreateObject("adodb.recordset")
rs3.open "select shopid from shop_action where  shop_action.username='"&username&"' and zhuangtai=8",conn,1,1
if rs3.recordcount>=4 then
  response.write "<script language=javascript>alert('对不起，您最多只能对比四件商品！');window.location.href='to.asp?action=show';</script>"
  response.end
else
  rs3.close
  set rs3=nothing
'//添加对比
  set rs=server.CreateObject("adodb.recordset")
  rs.open "select shopid,username,zhuangtai,zonger from shop_action",conn,1,3
  rs.addnew
  rs("shopid")		= shopid
  rs("username")	= username
  rs("zhuangtai")	= 8
  rs("zonger")		= 0
  rs.update
  rs.close
  response.Redirect "to.asp?action=show"
  set rs=nothing
end if

case "show"
%>

<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF">	
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td ><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td background="images/leftbg.jpg"><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg"></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
          <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <a href=index.asp><%=webname%></a> >> 商品对比 </td>
      </tr>
	  <tr>
        <td ></td>
      </tr>
      <tr>
        <td><br>
 <form name='form1' method='post' action="stowcart.asp">
                            <%

set rs=server.CreateObject("adodb.recordset")
rs.open "select shop_action.actionid,shop_action.shopid,products.bookname,products.shichangjia,products.huiyuanjia,products.vipjia,products.dazhe from products inner join  shop_action on products.bookid=shop_action.shopid where   shop_action.username='"&username&"' and  shop_action.zhuangtai=8",conn,1,1 
%>
                            <table width="170" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <%
			 do while not rs.eof
			 %>
                                <td align="center" valign="top"> 
                                  <%
				set rs2=server.CreateObject("adodb.recordset")
			   rs2.open "select * from products where bookid = "&rs("shopid"),conn,1,3
			   if rs2.EOF or rs2.bof  then
response.write"对不起，该商品不存在或已经删除"
else

shopname=rs2("bookname")
zhuang=rs2("zhuang")
shoppic=rs2("bookpic")
shichangjia=rs2("shichangjia")
huiyuanjia=rs2("huiyuanjia")
vipjia=rs2("vipjia")
kucun=rs2("kucun")
isbn=rs2("isbn")

yeshu=rs2("yeshu")

shopid=rs2("bookid")

hot=rs2("liulancount")
xiao=rs2("chengjiaocount")%>
                                  
                    <table width="170" align="center" cellpadding="0" cellspacing="0" bordercolor="#6699FF">
                      <tr> 
                        <td height="<%=upload14+10%>"> 
                          <%if shoppic="" then%>
                          <div align="left"> 
                            <table width="120" height="120" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td align="left" bgcolor="#FFFFFF" style="padding:5px"><img src=images/emptybook.gif width="75" height="75" border=0></td>
                              </tr>
                            </table>
                          </div>
                          <%else%>
                          <div align="left"> 
                            <table width="120" height="120" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td align="left" bgcolor="#FFFFFF" style="border:1px solid #CCCCCC;padding:5px"><img src=<%=shoppic%> width="100" height="100" border=0></td>
                              </tr>
                            </table>
                          </div>
                          <%end if%>                        </td>
                      </tr>
                      <tr> 
                        <td align="left"><font color="#ff6600"><a href="products.asp?id=<%=shopid%>" title="<%=shopname%>"> 
                          <%if len(shopname)>15 then
	response.write left(shopname,15)&""
	else
	response.write trim(shopname)
	end if%>
                          </a></font></td>
                      </tr>
                      <tr> 
                        <td align="left"><strong>市场价：</strong><s>￥<%=formatnumber(shichangjia,2)%>元</s></td>
                      </tr>
                      <tr> 
                        <td align="left"><font color="#FF0000"><font color="#FF0000"><strong>会员价：</strong>￥<%=formatnumber(huiyuanjia,2)%>元( 
                          <%response.write round(huiyuanjia/shichangjia,2)*10&"折"%>
                          )</font></font></td>
                      </tr>
                       
                      <tr> 
                        <td align="left"><font color="#FF0000"><font color="#FF0000"><strong>VIP 
                          价：</strong> 
                          <%if Request.Cookies("cnhww")("reglx")=2 then%>
                          ￥<%=trim(rs("Vipjia"))%>元 
                          <%else%>
                          <font color="#FF3333">VIP会员可见..</font> 
                          <%end if%>
                          </font></font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td align="left">【商品编号】<%=trim(rs2("grade"))%> </td>
                      </tr>
                      <tr> 
                        <td align="left">【商品规格】<%=trim(rs2("isbn"))%></td>
                      </tr>
                      <tr> 
                        <td align="left">【商品品牌】<%=trim(rs2("pingpai"))%></td>
                      </tr>
                      <tr> 
                        <td align="left">【赠送积分】<%=rs2("yeshu")%> 分</td>
                      </tr>
                      <tr> 
                        <td height="6" align="left">【商品数量】<%=trim(rs2("kucun"))%> 
                          <%=rs2("bookchuban")%> 
                          <%if rs2("kucun")>0 then%>
                          <font color="#999999">热卖中</font> 
                          <%else%>
                          <font color="#999999">缺货</font> 
                          <%end if%>                        </td>
                      </tr>
                      <tr> 
                        <td height="22" align="left">【商品折扣】 
                          <%response.write left(rs2("dazhe")*10,3)&"折"%>                        </td>
                      </tr>
                      <tr> 
                        <td height="11" align="left">【浏览次数】<%=trim(rs2("liulancount"))%>次</td>
                      </tr>
                      
                     
                    
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td align="left">【评论数目】 
                          <% idd=rs2("bookid")
		set rst=server.createobject("adodb.recordset")
		rst.open "select  * from review where bookid="&idd&" and shenhe=1 ",conn,1,1
		j=rst.recordcount
		
		%>
                         <a href="reviewlist.asp?id=<%=rs2("bookid")%>"> <%=j%>条评论 </a></td>
                      </tr>
                      <tr> 
                        <td align="left">【赠送积分】<%=yeshu%>分</td>
                      </tr>
                      <tr> 
                        <td align="left"><input name="newsshop" type="checkbox" id="newsshop" value="1" <%if rs2("newsbook")=1 then%>checked<%end if%> DISABLED>
                          新品 
                          <input name="bestshop" type="checkbox" id="newsshop" value="1" <%if rs2("bestbook")=1 then%>checked<%end if%> DISABLED>
                          推荐</td>
                      </tr>
                      <tr> 
                        <td align="left"><input name="tejiashop" type="checkbox" id="tejiashop" value="1" <%if rs2("tejiabook")=1 then%>checked<%end if%> DISABLED>
                          特价 </td>
                      </tr>
                      <tr> 
                        <td align="left"><input name=bookid type=checkbox checked value=<%=rs("shopid")%> > 
                          <a href=to.asp?action=del&actionid=<%=rs("actionid")%>><img src=images/trash.gif width=18 height=18 border=0></a></td>
                      </tr>
                      <tr> 
                        <td></td>
                      </tr>
                    </table></td>
                                <%end if
rs.movenext
loop
rs.close
set rs=nothing
%>
                              </tr>
                            </table>
                            <br>
                            <br><div align=center>
                            <input type="submit" name="Submit"  class=go-wenbenkuang value="放入购物车"/>
                            <input type="button" name="Submit"   class=go-wenbenkuang value="返回继续购物" onClick="javascript:window.history.go(-2)"/></div>
            </form>
                          <p> 
                            <script language="Javascript">
function closeinfo(){window.close();}
setTimeout("closeinfo()", 600000);
            </script>
                            <%
end select%>
          </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" height="40"></td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>