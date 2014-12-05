<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--Help center</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/ds_07.jpg); text-align:center;">
<!--#include file="include/head.asp"-->
<%
dim action
action=request.QueryString("action")
if InStr(action,"'")>0 then
response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
response.end
end if
%>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td ><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" ></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td align="center"><img src="images/body/left06.gif"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href="help.asp?action=gouwuliucheng">Shopping process</a> </td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href="help.asp?action=fukuan">Mode of payment</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=tiaokuan>The terms of the deal</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=yunshushuoming>Delivery notice</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=baomi>Security</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=shiyongfalv>copyright notice</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=lxwm>Contact us</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=about>About us</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;<a href=help.asp?action=shouhoufuwu>sales / service</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=gongzuoshijian>office hours</a></td>
                  </tr>
                  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=baozhuangbiaozhun>Individual packaging</a></td>
                  </tr>
				  <tr>
                    <td height="1" background="images/leftbg.jpg"></td>
                  </tr>
                  <tr>
                    <td height="24" background="images/leftbg.jpg">　<img src="images/body/dot_off.gif">&nbsp;&nbsp;<a href=help.asp?action=baozhuangshuoming>Packing specification</a></td>
                  </tr>
                  <tr>
                    <td height="24"><img src="images/leftendbg.jpg" width="190" height="36"></td>
                  </tr>
              </table></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" id="AutoNumber3" height="0" width="100%">
      <tr>
        <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="165" valign="top" bgcolor="#FFFFFF"><%select case action
		  case "fukuan"
		  %>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> The mode of payment </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select huikuanfangshi from webinfo",conn,1,1
				response.write trim(rs("huikuanfangshi"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "gouwuliucheng"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> shopping process</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gouwuliucheng from webinfo",conn,1,1
				response.write trim(rs("gouwuliucheng"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                  <%case "feiyong"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Delivery and cost </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table1">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select songhuofeiyong from webinfo",conn,1,1
				response.write trim(rs("songhuofeiyong"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "yunshushuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Delivery notice</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table2">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select yunshushuoming from webinfo",conn,1,1
				response.write trim(rs("yunshushuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "tiaokuan"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> The terms of the deal</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table3">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select jiaoyitiaokuan from webinfo",conn,1,1
				response.write trim(rs("jiaoyitiaokuan"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "yunshushuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Transportation instructions
</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table4">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select yunshushuoming from webinfo",conn,1,1
				response.write trim(rs("yunshushuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "gongzuoshijian"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >><a href=help.asp?action=gongzuoshijian>office hours</a></td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table5">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gongzuoshijian from webinfo",conn,1,1
				response.write trim(rs("gongzuoshijian"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "about"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >><a href=help.asp?action=about>About us</a></td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table5">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select about from webinfo",conn,1,1
				response.write trim(rs("about"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "shouhoufuwu"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Product sales / customer service </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table6">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select shouhoufuwu from webinfo",conn,1,1
				response.write trim(rs("shouhoufuwu"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "baomi"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Security</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table7">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baomi from webinfo",conn,1,1
				response.write trim(rs("baomi"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "lxwm"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Contact us</td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table8">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select lxwm from webinfo",conn,1,1
				response.write trim(rs("lxwm"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case "shiyongfalv"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>copyright notice </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select shiyongfalv from webinfo",conn,1,1
				response.write trim(rs("shiyongfalv"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "baozhuangbiaozhun"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Individual packaging standard </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baozhuangbiaozhun from webinfo",conn,1,1
				response.write trim(rs("baozhuangbiaozhun"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
				  <%case "baozhuangshuoming"%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>Packing specification </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select baozhuangshuoming from webinfo",conn,1,1
				response.write trim(rs("baozhuangshuoming"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%case else%>
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >>shopping process </td>
                    </tr>
                  </table>
                  <br>
                  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0" id="Table9">
                    <tr>
                      <td><%set rs=server.CreateObject("adodb.recordset")
				rs.open "select gouwuliucheng from webinfo",conn,1,1
				response.write trim(rs("gouwuliucheng"))
				rs.close
				set rs=nothing%>                      </td>
                    </tr>
                  </table>
                <%end select%>              </td>
            </tr>
        </table></td>
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
