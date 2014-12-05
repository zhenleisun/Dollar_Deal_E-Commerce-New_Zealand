<html><head><title>商城管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

Dim theInstalledObjects(24)
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    theInstalledObjects(16) = "JMail.SmtpMail"
    theInstalledObjects(17) = "CDONTS.NewMail"
    theInstalledObjects(18) = "Persits.MailSender"
    theInstalledObjects(19) = "SMTPsvg.Mailer"
    theInstalledObjects(20) = "DkQmail.Qmail"
    theInstalledObjects(21) = "Geocel.Mailer"
    theInstalledObjects(22) = "IISmail.Iismail.1"
'检查组件是否被支持
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

'检查组件版本
Function getver(Classstr)
On Error Resume Next
getver=""
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(Classstr)
If 0 = Err Then getver=xtestobj.version
Set xTestObj = Nothing
Err = 0
End Function
%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="2" align="center" background="../images/admin_bg_1.gif" height="25"><b><font color="#ffffff">服务器信息统计</font></b></td>
</tr>
<tr >
  <td colspan="2">&nbsp;</td>
  </tr>
<tr >
<td width="50%">&nbsp;服务器类型：<font face="Verdana"><%=Request.ServerVariables("OS")%> 
（IP：</font><%=Request.ServerVariables("LOCAL_ADDR")%>）</td>
<td width="50%" >
&nbsp;<%
      Response.Write theInstalledObjects(9)
      If Not IsObjInstalled(theInstalledObjects(9)) Then 
        Response.Write "<font color=red><b> ×</b></font>"
      Else
        Response.Write getver(theInstalledObjects(9)) & "<font color=#888888>(FSO 文本文件读写)</b><font color=green><b> √</b></font> "
      End If
      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;返回服务器的主机名、DNS别名或IP地址：<%=Request.ServerVariables("SERVER_NAME")%></td>
<td width="50%">
&nbsp;<%
      Response.Write theInstalledObjects(10)
   	  Response.Write "<font color=#888888>(ACCESS 数据库)</font>"
      If Not IsObjInstalled(theInstalledObjects(10)) Then 
        Response.Write "<font color=red><b> ×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(10))
      End If
      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
<td width="50%">
&nbsp;<%
      Response.Write theInstalledObjects(16)
		Response.Write "<font color=#888888>(w3 Jmail 收发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(16)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(16))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;脚本超时时间：<%=Server.ScriptTimeout%> 秒</td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(17)
	  Response.Write "<font color=#888888>(WIN虚拟SMTP 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(17)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(17))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;服务器端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(18)
	  Response.Write "<font color=#888888>(ASPemail 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(18)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(18))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;站点物理路径：<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(19)
	  Response.Write "<font color=#888888>(ASPemail 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(19)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(19))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;服务器IIS版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(20)
	  Response.Write "<font color=#888888>(dkQmail 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(20)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(20))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;服务器CPU个数：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(21)
	  Response.Write "<font color=#888888>(Geocel 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(21)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(21))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;服务器时间：<%=now%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(22)
	  Response.Write "<font color=#888888>(IISemail 发信)</font>"

      If Not IsObjInstalled(theInstalledObjects(22)) Then 
        Response.Write "&nbsp;<font color=red><b>×</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>√</b></font> " & getver(theInstalledObjects(22))
      End If

      Response.Write vbCrLf
%>
</td>
</tr>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">搜 索 用 户</font></b></td>
</tr>
<tr> 
<td height="50" > 
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<form name="form2" method="post" action="manageuser.asp?action=select">
<td>
<input name="namekey" type="text" id="namekey" onFocus="this.value=''" value="请输入关键字">
<input name="checkbox" type="checkbox" id="checkbox" value="1" checked>
模糊查询&nbsp;&nbsp;&nbsp;
<input type="submit" name="Submit2" value=" 开始查询 ">
</td>
</form>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="tableBorder">
  <tr> 
    <td class="forumRowHighlight" height=11>版权声明</td>
    <td class="forumRow" style="LINE-HEIGHT: 150%">　1、本软件为商业软件；本软件受国际版权法保护，任何单位或个人对本软件进行盗版复制，必追究其法律责任!<br>
      　2、<a href="http://www.cnhww.com/copyright.asp" target="_blank">软件著作权证书登记号:2007SR02047 
      软著登字第068042号</a> <BR>
      　2、您可以对本系统进行修改和美化，但必须保留完整的版权信息，不得将修改后的版本用于任何商业目的；<BR>
      　3、本软件受中华人民共和国《著作权法》《计算机软件保护条例》等相关法律、法规保护，作者保留一切权利。<BR>
      　4、如有可能，请在您的网站上添加本站链接,谢谢！ </td>
  </tr>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">搜 索 商 品</font></b></td>
</tr>
<tr > 
<form name="form2" method="post" action="manageproduct.asp">
<td height="50"> 
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td>
<input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关键字">
<select name="selectm" id="selectm">
<option value="bookname">按商品名称</option>
<option value="bookcontent">按商品说明</option>
<option value="bookid">按商品序号</option>
<option value="0">全部商品</option>
<option value="news">按新品</option>
<option value="tejia">按特价</option>
<option value="tuijian">按推荐</option>
</select>
<input type="submit" name="Submit2" value=" 开始查询 ">
</td>
</tr>
<tr>
<td height=40>除查询“所有商品”外，必须要输入关键字，才能查询<br>根据商品序号来查询时，关键字里只能输入数字。</td></tr>
</table>
</td>
</form>
</tr>
</table>
<!--#include file="foot.asp"-->
</body></html>