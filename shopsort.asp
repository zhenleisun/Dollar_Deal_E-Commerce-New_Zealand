<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--导航分类</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<meta name="keywords" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<link href="images/css.css" rel="stylesheet" type="text/css">
<!--#include file="include/head.asp"-->
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<TABLE width=996 border=0 align="center" cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR>
      <TD class=b vAlign=top align=left><table width="100%" height="49"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><br>
              <%
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select anclass,anclassid from bsort order by anclassidorder",conn,1,1
		  do while not rs.eof %>
              <table width="100%" height="70"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><TABLE width=980 height="53" border=0 align="center" cellPadding=0 cellSpacing=0>
                    <TBODY>
                      <TR>
                        <TD vAlign=top width=4 height=4><IMG height=4 
            src="images/new_line_004.gif" width=4></TD>
                        <TD background=images/new_line_008.gif height=4></TD>
                        <TD vAlign=top width=4 height=4><IMG height=4 
            src="images/new_line_005.gif" width=4></TD>
                      </TR>
                      <TR>
                        <TD width="1" rowspan="2" background=images/new_line_009.gif></TD>
                          <TD height="19"><img src="images/body/orange-bullet.gif" align="middle"> 
                            <%response.write "<br>&nbsp;&nbsp;&nbsp;<A href=class.asp?lx=big&anid="&rs("anclassid")&" class=top><b><font color=#FF3300>"&trim(rs("anclass"))&"</font></b></a>&nbsp;&nbsp;<br>"%>
                          </TD>
                        <TD width="1" rowspan="2" background=images/new_line_010.gif></TD>
                      </TR>
                      <TR>
                        <TD height="26"><% set rs2=server.CreateObject("adodb.recordset")
			rs2.open "select nclass,nclassid from ssort where anclassid="&rs("anclassid")&" order by nclassidorder",conn,1,1
			do while not rs2.eof
			response.write "&nbsp;&nbsp;&nbsp;<A href=class.asp?lx=small&anid="&rs("anclassid")&"&nid="&rs2("nclassid")&"><font color=#0099cc>"&trim(rs2("nclass"))&"</font></A>  " 
             rs2.movenext
			 loop
			 rs2.close
			 set rs2=nothing
			 response.write "<br>"%></TD>
                      </TR>
                      <TR>
                        <TD vAlign=top width=4 height=4><IMG height=4 
            src="images/new_line_006.gif" width=4></TD>
                        <TD background=images/new_line_011.gif></TD>
                        <TD vAlign=top width=4 height=4><IMG height=4 
            src="images/new_line_007.gif" width=4></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
                </tr>
              </table><br>
            <% rs.movenext
			loop
				rs.close
				set rs=nothing%></td>
          </tr>
        </table> 
		<br>       
      </td>
    </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>