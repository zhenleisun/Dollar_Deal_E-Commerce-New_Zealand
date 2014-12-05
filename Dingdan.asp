<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--Order center</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
  <LINK href="images/css.css" type=text/css rel=stylesheet></head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/ds_07.jpg); text-align:center;" >
<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="190" valign="top">
       <table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="userinfo.asp"--></td>
          </tr>
      </table>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td> <!--#include file="searc.asp"--></td>
          </tr>
		  <tr>
            <td height="15"><img src="images/leftendbg.jpg" width="190" height="36"></td>
          </tr>
        </table>    </td>
    <td width="771" valign="top"><DIV class=banner_mainM1>
        
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr> 
                  <td height=28 align="left"  background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    <a href=index.asp><%=webname%></a> >> Order center</td>
                </tr>
        </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
          <tr>
            <td valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
               <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
                <tr>
                    <td align=left><%  if request("dan")="" then %>
            <table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
              <tr align="center"> 
                <td bgcolor="f1f1f1"><b>Order inquiry</b></td>
              </tr>
              <form name="form" method="post" action="dingdan.asp">
                <tr bgcolor="#FFFFFF"> 
                  <td align="center" style='PADDING-LEFT: 5px'>Please enter your query Number： 
                    <input type="text" class="wenbenkuang" name="dan">
                    <input type="submit" name="Submit" value="Order inquiry">                  </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'>&nbsp;</td>
                </tr>
              </form>
            </table><%end if %>
            <br>
            <%dim dingdan
			dingdan=request("dan")
			if dingdan<>"" then
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select products.bookid,products.shjiaid,products.bookname,products.shichangjia,products.huiyuanjia,orders.actiondate,orders.shousex,orders.danjia,orders.feiyong,orders.fapiao,orders.userzhenshiname,orders.shouhuoname,orders.dingdan,orders.youbian,orders.liuyan,orders.zhifufangshi,orders.songhuofangshi,orders.zhuangtai,orders.zonger,orders.useremail,orders.usertel,orders.shouhuodizhi,orders.bookcount from products inner join orders on products.bookid=orders.bookid where dingdan='"&dingdan&"' ",conn,1,1
			if rs.eof and rs.bof then
			response.write "<script>alert('Without this order, please input again！');javascript:history.go(-1);</script>"
			response.end
			end if
			%>
            <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="10">
                  <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                    <tr align="center" bgcolor="f1f1f1"> 
                      <td><strong>Name</strong></td>
                      <td><strong>Order quantity</strong></td>
                      <td><strong>Member price</strong></td>
                      <td><strong>Subtotal amount</strong></td>
                  </tr>
                  <%zongji=0
		do while not rs.eof%>
                <tr bgcolor="#FFFFFF">
                  <td align="center" style='PADDING-LEFT: 5px'><a href=products.asp?id=<%=rs("bookid")%> ><%=trim(rs("bookname"))%></a></td>
                    <td align="center"><%=rs("bookcount")%></td>
                    <td align="center">$ <%=rs("danjia")&""%></td>
                    <td align="center">$ <%=rs("zonger")&""%></td>
                  </tr>
                <%zongji=rs("zonger")+zongji
		feiyong=rs("feiyong")
		rs.movenext
		loop
		rs.movefirst%>
         <tr bgcolor="#FFFFFF">
          <td colspan="4"align="right">The total order：$<%=zongji%>＋Cost( $<%=feiyong%>)　　Total：$ <font color="#ff0000">
		  <%=formatnumber(zongji+feiyong,2)%></font> 
                      &nbsp;&nbsp;&nbsp;&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
			<tr>
              <td>&nbsp;</td>
            </tr>
            <tr>
              <td>
                  <table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#cccccc" align="center">
                    <tr bgcolor="f1f1f1"> 
                      <td colspan="2" align="left">　<strong>You this order number: [<%=dingdan%>] detailed information is as follows：</strong></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="21%" align="right">Order status：</td>
                      <td width="79%" style="PADDING-LEFT: 12px"> <%call zhuang()%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Name of consignee：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("userzhenshiname"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">The delivery address：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("shouhuodizhi"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Methods of payment：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Telephone：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("usertel"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Email：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("useremail"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">The delivery method：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Zip code：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("youbian"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">If want to invoice：</td>
                      <td align="left" style="PADDING-LEFT: 12px"> 
                        <%if rs("fapiao")=1 then %>
                        <font color=red>You need the invoice</font> 
                        <%else%>
                        Don't need invoice 
                        <%end if%>                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Your message：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=trim(rs("liuyan"))%></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td align="right">Place an order date：</td>
                      <td align="left" style="PADDING-LEFT: 12px"><%=rs("actiondate")%></td>
                    </tr>
                  </table>                </td>
              </tr>
            </table>
            <br>
          <%end if
			sub zhuang()
			select case rs("zhuangtai")
			case "1"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          Without any treatment<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="2" DISABLED>
          The user has to draw money<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          Service providers have received<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          Service providers have delivery<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          The user has received the goods
          <%else%>
          <input name="zhuangtai" type="checkbox" id="zhuangtai" value="3">
          The order has been confirmed<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          The orders have been completed
          <%end if%>
          <%case "2"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          Without any treatment<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          The user has to draw money<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" DISABLED>
          Service providers have received<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          Service providers have delivery<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          The user has received the goods
          <%case "3"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          Without any treatment<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          The user has to draw money<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          Service providers have received<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" DISABLED>
          Service providers have delivery<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" DISABLED>
          The user has received the goods
          <%else%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          The order has been confirmed<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox1" value="checkbox" >
          The orders have been completed
          <%end if%>
          <%case "4"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          Without any treatment<span style='font-family:Wingdings;'>à</span>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          The user has to draw money<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          Service providers have received<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          Service providers have delivery<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="zhuangtai" value="5" >
          The user has received the goods
          <%case "5"%>
          <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>
          Without any treatment<span style='font-family:Wingdings;'>à</span>
          <%if rs("zhifufangshi")<>4 then %>
          <input name="checkbox" type="checkbox" id="zhuangtai" value="2" checked DISABLED>
          The user has to draw money<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>
          Service providers have received<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>
          Service providers have delivery<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          The user has received the goods
          <%else%>
          <input name="checkbox3" type="checkbox" DISABLED value="checkbox" checked>
          The order has been confirmed<span style='font-family:Wingdings;'>à</span>
          <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>
          The orders have been completed
          <%end if%>
          <%end select
end sub%>
          <br>&nbsp;</td>
                 </tr>
              </table>              </td>
          </tr>
        </table>
    </DIV></td>
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