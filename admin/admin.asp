<html><head><title>�̳ǹ���ϵͳ</title>
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
'�������Ƿ�֧��
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

'�������汾
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
<td colspan="2" align="center" background="../images/admin_bg_1.gif" height="25"><b><font color="#ffffff">��������Ϣͳ��</font></b></td>
</tr>
<tr >
  <td colspan="2">&nbsp;</td>
  </tr>
<tr >
<td width="50%">&nbsp;���������ͣ�<font face="Verdana"><%=Request.ServerVariables("OS")%> 
��IP��</font><%=Request.ServerVariables("LOCAL_ADDR")%>��</td>
<td width="50%" >
&nbsp;<%
      Response.Write theInstalledObjects(9)
      If Not IsObjInstalled(theInstalledObjects(9)) Then 
        Response.Write "<font color=red><b> ��</b></font>"
      Else
        Response.Write getver(theInstalledObjects(9)) & "<font color=#888888>(FSO �ı��ļ���д)</b><font color=green><b> ��</b></font> "
      End If
      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;���ط���������������DNS������IP��ַ��<%=Request.ServerVariables("SERVER_NAME")%></td>
<td width="50%">
&nbsp;<%
      Response.Write theInstalledObjects(10)
   	  Response.Write "<font color=#888888>(ACCESS ���ݿ�)</font>"
      If Not IsObjInstalled(theInstalledObjects(10)) Then 
        Response.Write "<font color=red><b> ��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(10))
      End If
      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
<td width="50%">
&nbsp;<%
      Response.Write theInstalledObjects(16)
		Response.Write "<font color=#888888>(w3 Jmail �շ���)</font>"

      If Not IsObjInstalled(theInstalledObjects(16)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(16))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;�ű���ʱʱ�䣺<%=Server.ScriptTimeout%> ��</td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(17)
	  Response.Write "<font color=#888888>(WIN����SMTP ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(17)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(17))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;�������˿ڣ�<%=Request.ServerVariables("SERVER_PORT")%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(18)
	  Response.Write "<font color=#888888>(ASPemail ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(18)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(18))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;վ������·����<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(19)
	  Response.Write "<font color=#888888>(ASPemail ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(19)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(19))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;������IIS�汾��<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(20)
	  Response.Write "<font color=#888888>(dkQmail ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(20)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(20))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;������CPU������<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(21)
	  Response.Write "<font color=#888888>(Geocel ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(21)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(21))
      End If

      Response.Write vbCrLf
%></td>
</tr>
<tr >
<td width="50%">&nbsp;������ʱ�䣺<%=now%></td>
<td width="50%">&nbsp;<%
      Response.Write theInstalledObjects(22)
	  Response.Write "<font color=#888888>(IISemail ����)</font>"

      If Not IsObjInstalled(theInstalledObjects(22)) Then 
        Response.Write "&nbsp;<font color=red><b>��</b></font>"
      Else
        Response.Write "&nbsp;<font color=green><b>��</b></font> " & getver(theInstalledObjects(22))
      End If

      Response.Write vbCrLf
%>
</td>
</tr>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�� �� �� ��</font></b></td>
</tr>
<tr> 
<td height="50" > 
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<form name="form2" method="post" action="manageuser.asp?action=select">
<td>
<input name="namekey" type="text" id="namekey" onFocus="this.value=''" value="������ؼ���">
<input name="checkbox" type="checkbox" id="checkbox" value="1" checked>
ģ����ѯ&nbsp;&nbsp;&nbsp;
<input type="submit" name="Submit2" value=" ��ʼ��ѯ ">
</td>
</form>
</tr>
</table>
</td>
</tr>
</table>
<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="tableBorder">
  <tr> 
    <td class="forumRowHighlight" height=11>��Ȩ����</td>
    <td class="forumRow" style="LINE-HEIGHT: 150%">��1�������Ϊ��ҵ�����������ܹ��ʰ�Ȩ���������κε�λ����˶Ա�������е��渴�ƣ���׷���䷨������!<br>
      ��2��<a href="http://www.cnhww.com/copyright.asp" target="_blank">�������Ȩ֤��ǼǺ�:2007SR02047 
      �������ֵ�068042��</a> <BR>
      ��2�������ԶԱ�ϵͳ�����޸ĺ������������뱣�������İ�Ȩ��Ϣ�����ý��޸ĺ�İ汾�����κ���ҵĿ�ģ�<BR>
      ��3����������л����񹲺͹�������Ȩ����������������������������ط��ɡ����汣�������߱���һ��Ȩ����<BR>
      ��4�����п��ܣ�����������վ����ӱ�վ����,лл�� </td>
  </tr>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�� �� �� Ʒ</font></b></td>
</tr>
<tr > 
<form name="form2" method="post" action="manageproduct.asp">
<td height="50"> 
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td>
<input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؼ���">
<select name="selectm" id="selectm">
<option value="bookname">����Ʒ����</option>
<option value="bookcontent">����Ʒ˵��</option>
<option value="bookid">����Ʒ���</option>
<option value="0">ȫ����Ʒ</option>
<option value="news">����Ʒ</option>
<option value="tejia">���ؼ�</option>
<option value="tuijian">���Ƽ�</option>
</select>
<input type="submit" name="Submit2" value=" ��ʼ��ѯ ">
</td>
</tr>
<tr>
<td height=40>����ѯ��������Ʒ���⣬����Ҫ����ؼ��֣����ܲ�ѯ<br>������Ʒ�������ѯʱ���ؼ�����ֻ���������֡�</td></tr>
</table>
</td>
</form>
</tr>
</table>
<!--#include file="foot.asp"-->
</body></html>