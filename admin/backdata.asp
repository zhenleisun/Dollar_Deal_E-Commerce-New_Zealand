<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
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
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�������ݿ�</font></b></td>
</tr>
<tr>
<td width="100%" >
<font color="#FF0000">ע�⣺����������ҪFSO���֧�֣�FSO�������ذ�����·�����ǳ���ռ��Ŀ¼�����·������������Щ�ռ䱸�ݺ�,�ڱ����ϲ�����ACCESS��,�˹���С��ʹ�á�</font></td>
</tr>
<tr><form method="post" action="Backdata.asp?action=Backup">
<td width="100%" >
��ǰ���ݿ�·����
        <input type=text size=24 name=DBpath value="../cnhwwdata/cnhww.asp"> 
����ȷ��д����ǰʹ�õ����ݿ�·����<BR>
�������ݿ�Ŀ¼��<input type=text size=24 name=bkfolder value=../Databackup>
        ���Ŀ¼�����ڣ������Զ�������<BR>
        �������ݿ����ƣ�Ĭ��Ϊshop.mdb ���ݺ��뵽��ӦĿ¼���������ݱ���! <br>
<input type=submit value="��ʼ����"></td>
<tr>
<td width="100%" >
��������д���ݿ�·�������ݿ��������ƣ������Ĭ�����ݿ��ļ�Ϊshop.mdb<br>
����������������������������ݿ⣬�Ա�֤���ݵİ�ȫ��<br></td>
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
response.write "���ݿⱸ����ɣ����������������<br>����ʹ�� FTP ���߽����ݿⱸ�ݣ��Ա�֤���ݰ�ȫ"
else
response.write "�Ҳ���������Ҫ���ݵ��ļ���"
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
