<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">ѹ�����ݿ� ( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ ) </font></b></td>
</tr>
<tr bgcolor="#fbf4f4">
<td width=100%>
<font color=red>ע�⣺�������ݿ��������·��,���������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ����ѡ�񱸷����ݿ����ѹ����</font></td>
</tr>
<tr>
<%
Const JET_3X = 4
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
     Set Engine = CreateObject("JRO.JetEngine")
      If boolIs97 = "True" Then
             Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
             "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
              & "Jet OLEDB:Engine Type=" & JET_3X
      Else
	Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
	"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
       End If
       fso.CopyFile strDBPath & "temp.mdb",dbpath
       fso.DeleteFile(strDBPath & "temp.mdb")
       Set fso = nothing
       Set Engine = nothing
       CompactDB = "����ָ�������ݿ�, " & dbpath & ", �Ѿ���ѹ��" & vbCrLf
Else
       CompactDB = "������������ݿ�·��������δ�ҵ���������" & vbCrLf
End If
End Function
%>
<form name="compress" method="post" action="compressdata.asp">
<td width=100% bgcolor=#fbf4f4>
ѹ�����ݿ⣺
  <input type="text" name="dbpath" size="20" value="../cnhwwdata/cnhww.asp">
<input type="submit" name="submit" value="��ʼѹ��">
</td>

<tr bgcolor="#fbf4f4">
<td width=100%>
<input type="checkbox" name="boolIs97" value="True">
���ʹ�� Access 97 ���ݿ���ѡ��Ĭ��Ϊ Access 2000 ���ݿ⣩
</td>
</tr>
</form>
<%
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")
If dbpath <> "" Then
       dbpath = server.mappath(dbpath)
       response.write(CompactDB(dbpath,boolIs97))
End If
%>
</td></tr></table>
<!--#include file="foot.asp"-->
</body>
</html>
