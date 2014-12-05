<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
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
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">压缩数据库 ( 需要FSO支持，FSO相关帮助请看微软网站 ) </font></b></td>
</tr>
<tr bgcolor="#fbf4f4">
<td width=100%>
<font color=red>注意：输入数据库所在相对路径,并输入数据库名称（正在使用中数据库不能压缩，选择备份数据库进行压缩）</font></td>
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
       CompactDB = "您所指定的数据库, " & dbpath & ", 已经被压缩" & vbCrLf
Else
       CompactDB = "您所输入的数据库路径或名称未找到，请重试" & vbCrLf
End If
End Function
%>
<form name="compress" method="post" action="compressdata.asp">
<td width=100% bgcolor=#fbf4f4>
压缩数据库：
  <input type="text" name="dbpath" size="20" value="../cnhwwdata/cnhww.asp">
<input type="submit" name="submit" value="开始压缩">
</td>

<tr bgcolor="#fbf4f4">
<td width=100%>
<input type="checkbox" name="boolIs97" value="True">
如果使用 Access 97 数据库请选择（默认为 Access 2000 数据库）
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
