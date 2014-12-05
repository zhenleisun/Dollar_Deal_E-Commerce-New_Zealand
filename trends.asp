<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="#">
<meta name="keywords" content="#">

<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/ds_07.jpg); text-align:center;">
<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" align="left" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
        <td><!--#include file="searc.asp"--></td>
      </tr>
	  <tr>
        <td background="images/leftbg.jpg"><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg"></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table cellspacing=0 cellpadding=0 width=100% align=center border=0>
      <tbody>
        <tr>
          <td class=b valign=top align=left width=100% ><%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""Illegal access!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""Illegal access!"");location.href=""index.asp"";</script>"
end if%>
              <%dim newsid
newsid=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news where newsid="&newsid,conn,1,3
rs("viewcount")=rs("viewcount")+1
rs.update
%>
              <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr>
                  <td width="100%" align="center" valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="left" height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Mall news</td>
                      </tr>
                      <tr bgcolor="#ffffff">
                        <td  valign="top"><table width="94%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td colspan="3" align="center"><br>
                                <font color=#F46404><strong><%=trim(rs("newsname"))%></strong></font><br>
                                <br>
                                Author：<%=trim(rs("addname"))%> Hits <font color=red><%=rs("viewcount")%></font> 
                                【Font size: <A onclick="Zoom.style.fontSize='19px';" href="#">Big</A> 
                                <A  onclick="Zoom.style.fontSize='16px';"  href="#"></A> <A    onclick="Zoom.style.fontSize='12px';"  href="#">small</A>】 Release time：<%=year(rs("adddate"))&"Year"&month(rs("adddate"))&"Month"&day(rs("adddate"))&"Day"%> 
                                <a class="b12" href="#" onclick="if (window.print) window.print();return false">Print this page</a> 
                                <hr width="720" size="1" noshade>
                                <font color=#F46404><b> </b></font></td>
                            </tr>
                            <tr>
                              <td colspan="3"><br><DIV id=Content><FONT id=Zoom><%=trim(rs("newscontent"))%></div><br></td>
                            </tr>
                            <tr>
                              <td colspan="3"></td>
                            </tr>
                            <tr>
                              <td colspan="3" align="center" ><br>
                              <table width="98%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
                                  </tr>
                                <tr>
                                  <td height="40">Author：<%=trim(rs("addname"))%> </td>
                                  <td>Release time：<%=year(rs("adddate"))&"Year"&month(rs("adddate"))&"Month"&day(rs("adddate"))&"Day"%></td>
                                  <td height="25">Has been viewed <font color=red><%=rs("viewcount")%></font> </td>
                                </tr>
                              </table></td></tr>
                            <tr>
                              <td colspan="3" align="right" ><a href='javascript:onclick=history.go(-1)'><img src="images/mingle/infbk.gif" width="108" height="24" border="0" ></a>&nbsp;&nbsp;&nbsp;</td>
                            </tr>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            <%rs.close
set rs=nothing%></td>
          <td ></td>
        </tr>
      </tbody>
    </table></td>
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
