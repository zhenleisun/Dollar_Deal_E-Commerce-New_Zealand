<%@ Language="VBScript" %>
<% Option Explicit %>
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
'��ʹ�������������ֱ�ӽ����н����ʾ�ڿͻ���
Response.Buffer = False
'�������������
Dim ObjTotest(26,4)
ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO �ı��ļ���д)"
ObjTotest(10,0) = "adodb.linksion"
	ObjTotest(10,1) = "(ADO ���ݶ���)"
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp �ļ��ϴ�)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans �ļ�����)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(�ļ��ϴ����)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload �ļ��ϴ�)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac �ļ��ϴ�)"
ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail �ʼ��շ�) <a href='http://www.5757.net'>�����ֲ�����</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(���� SMTP ����)"
ObjTotest(18,0) = "Persits.MailSender"
	ObjTotest(18,1) = "(ASPemail ����)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail ����)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail ����)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel ����)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail ����)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail ����)"
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA ��ͼ���д���)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac ��ͼ���д���)"
public IsObj,VerObj
'���Ԥ�����֧��������汾
dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	dim TestObj
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		'��л���ѵı�����
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next
'�������Ƿ�֧�ּ�����汾���ӳ���
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	dim TestObj
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'��л����5757�ı�����
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">����������</font></b></td>
</tr>
<tr> 
<td align="center" >
		<table width=500 border=1 cellpadding=0 cellspacing=0 style="border-collapse: collapse" bordercolor="#990000">
          <tr  height=18>
            <td align=left>&nbsp;��������</td>
            <td>&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;������IP</td>
            <td>&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;�������˿�</td>
            <td>&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;������ʱ��</td>
            <td>&nbsp;<%=now%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;IIS�汾</td>
            <td>&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;�ű���ʱʱ��</td>
            <td>&nbsp;<%=Server.ScriptTimeout%> ��</td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;���ļ�·��</td>
            <td>&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;������CPU����</td>
            <td>&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;��������������</td>
            <td>&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
          </tr>
          <tr  height=18>
            <td align=left>&nbsp;����������ϵͳ</td>
            <td>&nbsp;<%=Request.ServerVariables("OS")%></td>
          </tr>
        </table>
		<%
	Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>��ָ��������ļ������"
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
	  Else
		Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����������汾�ǣ�" & VerObj & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%>
		�� IIS�Դ���ASP��� ��
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="500">
          <tr height=18 class=backs align=center> 
            <td width=320>�� �� �� ��</td>
            <td width=130>֧�ּ��汾</td>
          </tr>
          <%For i=0 to 10%>
          <tr height="18" class=backq> 
            <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
            <td align=left>&nbsp; 
              <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%>
            </td>
          </tr>
          <%next%>
        </table>
        �� �������ļ��ϴ��͹������ ��  
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="500">
          <tr height=18 class=backs align=center> 
            <td width=320>�� �� �� ��</td>
            <td width=130>֧�ּ��汾</td>
          </tr>
          <%For i=11 to 15%>
          <tr height="18" class=backq> 
            <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
            <td align=left>&nbsp; 
              <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%>
            </td>
          </tr>
          <%next%>
        </table>
        �� �������շ��ʼ���� �� 
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="500">
          <tr height=18 class=backs align=center> 
            <td width=320>�� �� �� ��</td>
            <td width=130>֧�ּ��汾</td>
          </tr>
          <%For i=16 to 23%>
          <tr height="18" class=backq> 
            <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
            <td align=left>&nbsp; 
              <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%>
            </td>
          </tr>
          <%next%>
        </table>
		�� ͼ������� ��
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="500">
          <tr height=18 class=backs align=center> 
            <td width=320>�� �� �� ��</td>
            <td width=130>֧�ּ��汾</td>
          </tr>
          <%For i=24 to 25%>
          <tr height="18" class=backq> 
            <td height="39" align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></td>
            <td align=left>&nbsp; 
              <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%>
            </td>
          </tr>
          <%next%>
        </table>
		�� �������֧�������� ��<br>
        ��������������������Ҫ���������ProgId��ClassId�� 
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="500">
          <form action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
            <tr height="18" class=backq> 
              <td align=center height=30> 
                <input class=input type=text value="" name="classname" size=40>
                <input type=submit value=" ȷ �� " class=backc id=submit1 name=submit1>
                <input type=reset value=" �� �� " class=backc id=reset1 name=reset1>
              </td>
            </tr>
          </form>
        </table>
		�� ASP�ű����ͺ������ٶȲ��� ��<br>
        �����÷�����ִ��50��Ρ�1��1���ļ��㣬��¼����ʹ�õ�ʱ�䡣 
        <table class=backq border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#990000" width="499">
          <tr height=18 class=backs align=center> 
            <td width=351>��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;��</td>
            <td width=142>���ʱ��</td>
          </tr>
          <tr height=18> 
            <td align=left width="351">&nbsp;�й�Ƶ������������2002-08-06 9:29��</td>
            <td width="142">&nbsp;610.9 ����</td>
          </tr>
          <tr height=18> 
            <td align=left width="351">&nbsp;��������west263������2002-08-06 9:29��</td>
            <td width="142">&nbsp;357.8 ����</td>
          </tr>
          <tr height=18> 
            <td align=left width="351">&nbsp;�����й�����������2002-08-06 9:29��</td>
            <td width="142">&nbsp;353.1 ����</td>
          </tr>
          <tr height=18> 
            <td align=left width="351">&nbsp;����Ƽ�tonydns������2002-10-13 14:19��</td>
            <td width="142">&nbsp;303.2 ����</td>
          </tr>
          <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
            <%
	dim t1,t2,lsabc,thetime
	t1=timer
	for i=1 to 500000
		lsabc= 1 + 1
	next
	t2=timer
	thetime=cstr(int(( (t2-t1)*10000 )+0.5)/10)
%>
            <tr height=18> 
              <td align=left width="351">&nbsp;<font color=red>������ʹ�õ���̨������</font>&nbsp;</td>
              <td width="142">&nbsp;<font color=red><%=thetime%> ����</font></td>
            </tr>
          </form>
        </table>
      </div>    </td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</BODY>
</HTML>