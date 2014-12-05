<!--#include file="conn.asp"-->
<%if request.cookies("Cnhww")("username")="" then
response.write "<script language=javascript>alert('对不起，您还没有登陆！');history.go(-1);</script>"
response.End
end if%>
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--密码找回</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/xxl_09.png); text-align:center;">
<!--#include file="include/head.asp"-->
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
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
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width=812 valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=request.cookies("Cnhww")("username")%> >> 订单详细信息</td>
      </tr>
      <tr>
        <td><br>
        <%dim dingdan
dingdan=request.QueryString("dan")
set rs=server.CreateObject("adodb.recordset")
rs.open "select products.bookid,products.shjiaid,products.bookname,products.shichangjia,products.huiyuanjia,orders.actiondate,orders.shousex,orders.danjia,orders.feiyong,orders.fapiao,orders.userzhenshiname,orders.shouhuoname,orders.dingdan,orders.youbian,orders.liuyan,orders.zhifufangshi,orders.songhuofangshi,orders.zhuangtai,orders.zonger,orders.useremail,orders.usertel,orders.shouhuodizhi,orders.bookcount from products inner join orders on products.bookid=orders.bookid where orders.username='"&request.cookies("Cnhww")("username")&"' and dingdan='"&dingdan&"' ",conn,1,1
if rs.eof and rs.bof then
response.write "<p align=center>此订单中有商品已被管理员删除，无法进行正确计算!<br>订单取消，请通知管理员或重新下订单!</p>"
response.End
end if
%>
          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="10"><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#f1f1f1">
                <tr align="center">
                    <td><strong>商品名称</strong></td>
                    <td><strong>订购数量</strong></td>
                    <td><strong>会员价格</strong></td>
                    <td><strong>金额小计</strong></td>
                  </tr>
                  <%zongji=0
		do while not rs.eof%>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'><a href=products.asp?id=<%=rs("bookid")%> ><%=trim(rs("bookname"))%></a></td>
                    <td align="center"><%=rs("bookcount")%></td>
                    <td align="center"><%=rs("danjia")&"元"%></td>
                    <td align="center"><%=rs("zonger")&"元"%></td>
                  </tr>
                <%zongji=rs("zonger")+zongji
		feiyong=rs("feiyong")
		rs.movenext
		loop
		rs.movefirst%>
         <tr bgcolor="#FFFFFF">
          <td colspan="4"align="right">订单总额：<%=zongji%>元＋费用(<%=feiyong%>元)　　共计：<font color="#ff0000"><%=zongji+feiyong%></font>元 
                      &nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#f1f1f1" align="center">
                <tr>
                  <td colspan="3">　<strong>您</font>此次订单号为：[ <%=dingdan%> ]</font></font> 详细信息如下：</font></strong></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td width="15%" align="right">订单状态：</td>
                      <td colspan="2"><table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
                        <tr>
                          <form name="form1" method="post" action="saveorder.asp?dan=<%=dingdan%>&action=save">
                            <td><%zhuang()%>
                              <br>
                              <%if rs("zhuangtai")<>5 then %>
                              <input class="go-wenbenkuang" name="submit" value="修改订单状态" type="submit">
                              <%else
					response.write "<font color=black><b>该订单交易工作结束!</b></font>"
					end if%>                              </td>
                            </form>
                          </tr>
                        </table></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">收货人姓名：</td>
                      <td width="47%" style="PADDING-LEFT: 12px"><%=trim(rs("userzhenshiname"))%></td>
                      <td width="38%" style="PADDING-LEFT: 12px"><img src="images/mingle/charge.gif" width="35" height="32"><a href="help.asp?action=feiyong" target="_blank" ><font color="2782B3">送货相关资费说明</font></a></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">收货地址：</td>
                      <td style="PADDING-LEFT: 12px"><%=trim(rs("shouhuodizhi"))%></td>
                      <td style="PADDING-LEFT: 12px"><img src="images/mingle/post.gif" width="38" height="32"><a href="help.asp?action=lxwm" target="_blank"><font color="2782B3">商家联系方式</font></a></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
				<td align="right">支付方式：</td>
                      <td style="PADDING-LEFT: 12px"><%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>                      </td>
				
                  
                      <td style="PADDING-LEFT: 12px"><img src="images/mingle/money.gif" width="44" height="32"><a href=help.asp?action=fukuan target="_blank"><font color="2782B3">商家汇款方式</font></a></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">联系电话：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%=trim(rs("usertel"))%></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">电子邮件：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%=trim(rs("useremail"))%></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">送货方式：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>                      </td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">邮编：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%=trim(rs("youbian"))%></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">是否要发票：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%if rs("fapiao")=1 then %>
                        <font color=red>需要发票</font>
                        <%else%>
                        不需要发票
                        <%end if%>                      </td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">您的留言：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%=trim(rs("liuyan"))%></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="right">下单日期：</td>
                      <td colspan="2" style="PADDING-LEFT: 12px"><%=rs("actiondate")%></td>
                    </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="40" colspan="3" align="center"><%if rs("zhuangtai")=1 then%>
                    <input class="go-wenbenkuang" type="submit" name="submit" value="删除订单" onClick="if(confirm('您确定要删除吗?')) location.href='saveorder.asp?action=del&dan=<%=dingdan%>';else return;">
                    <%end if%>
                    <input class=go-wenbenkuang onClick=javascript:window.close() type=reset value=关闭窗口 name=submit>                    </td>
                    </tr>
                </table></td>
              </tr>
            </table>
            <br>
          <%
sub zhuang()
select case rs("zhuangtai")
case "1"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="2" checked>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%else%>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="3">
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          订单已完成
          <%end if%>
          <%case "2"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%case "3"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          用户已经收到货
          <%else%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox1" value="checkbox" >
          订单已完成
          <%end if%>
          <%case "4"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="zhuangtai" value="5" >
          用户已经收到货
          <%case "5"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          未作任何处理<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          用户已经划出款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          服务商已经收到款<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          服务商已经发货<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          用户已经收到货
          <%else%>
          <input name="checkbox3" type="checkbox" DISABLED value="checkbox" checked>
          订单已确认<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          订单已完成
          <%end if%>
          <%end select
end sub%>
          <br></td></tr>
    </table></td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
