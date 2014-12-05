<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
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

<%
dim bkfolder
dim bkdbname
dim fso
dim folderpath,fso1,f
call main()
conn.close
set conn=nothing
sub main()
%>
<%
if request("action")="Backup" then
call backupdata()
else
%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">备份数据库</font></b></td>
</tr>
<tr>
<td width="100%" >
<font color="#FF0000">注意：备份数据需要FSO组件支持，FSO组件的相关帮助！路径都是程序空间根目录的相对路径！可能在有些空间备份后,在本机上不能用ACCESS打开,此功能小心使用。</font></td>
</tr>
<tr><form method="post" action="Backdata.asp?action=Backup">
<td width="100%" >
当前数据库路径：
        <input type=text size=24 name=DBpath value="../cnhwwdata/cnhww.asp"> 
请正确添写您当前使用的数据库路径！<BR>
备份数据库目录：<input type=text size=24 name=bkfolder value=../Databackup>
        如果目录不存在，程序将自动创建！<BR>
        备份数据库名称：默认为shop.mdb 备份后请到相应目录下下载数据备份! <br>
<input type=submit value="开始备份"></td>
<tr>
<td width="100%" >
在上面填写数据库路径及数据库完整名称，程序的默认数据库文件为shop.mdb<br>
您可以用这个功能来备份您的数据库，以保证数据的安全！<br></td>
</tr></form>
</table>
<%
end if
%>
<%
end sub
sub backupdata()
Dbpath=request.form("Dbpath")
Dbpath=server.mappath(Dbpath)
bkfolder=request.form("bkfolder")
bkdbname="shop.mdb"
Set Fso=server.createobject("scripting.filesystemobject")
if fso.fileexists(dbpath) then
If CheckDir(bkfolder) = True Then
fso.copyfile dbpath,bkfolder& "\"& bkdbname
else
MakeNewsDir bkfolder
fso.copyfile dbpath,bkfolder& "\"& bkdbname
end if
response.write "数据库备份完成，请进行其他操作！<br>建议使用 FTP 工具将数据库备份，以保证数据安全"
else
response.write "找不到您所需要备份的文件！"
end if
end sub
Function CheckDir(FolderPath)
folderpath=Server.MapPath(".")&"\"&folderpath
Set fso1 = CreateObject("Scripting.FileSystemObject")
If fso1.FolderExists(FolderPath) then
CheckDir = True
Else
CheckDir = False
End if
Set fso1 = nothing
End Function
Function MakeNewsDir(foldername)
Set fso1 = CreateObject("Scripting.FileSystemObject")
Set f = fso1.CreateFolder(foldername)
MakeNewsDir = True
Set fso1 = nothing
End Function
%>
<!--#include file="foot.asp"-->
</body>
</html>
