<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='index.asp';</script>"
response.End
else
if session("flag")>2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<!--#include file="upload.inc"-->
<%
	Set Upload = New UpFile_Class						
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png"		
	Upload.MaxSize = 10240000	
	Upload.GetDate()	
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 2 : Response.Write "ͼƬ��С���������� "&Dvbbs.Forum_Setting(56)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 3 : Response.Write "���ϴ����Ͳ���ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		End Select
	Else
		 FormPath=Upload.Form("filepath")
		 For Each FormName in Upload.file		
			 Set File = Upload.File(FormName)	
			 If File.Filesize<10 Then
		 		Response.Write "����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			End If
 			FileName=FormPath&UserFaceName(FileExt)
 			If File.FileSize>0 Then   
				File.SaveToFile Server.mappath(FileName)   
				response.write "<script>window.opener.document."&upload.form("FormName")&"."&upload.form("EditName")&".value='"&FileName&"'</script>"
				Response.Write "<script language=""javascript"">window.alert(""�ļ��ϴ��ɹ�!�벻Ҫ�޸����ɵ����ӵ�ַ��"");window.close();</script>"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
Private Function CheckFileExt(FileExt)
	Dim ForumUpload,i
	ForumUpload="gif,jpg,bmp,jpeg,png"
	ForumUpload=Split(ForumUpload,",")
	CheckFileExt=False
	For i=0 to UBound(ForumUpload)
		If LCase(FileExt)=Lcase(Trim(ForumUpload(i))) Then
			CheckFileExt=True
			Exit Function
		End If
	Next
End Function
Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Lcase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"asp","")
	FixName = Replace(FixName,"asa","")
	FixName = Replace(FixName,"aspx","")
	FixName = Replace(FixName,"cer","")
	FixName = Replace(FixName,"cdx","")
	FixName = Replace(FixName,"htr","")
End Function
Private Function UserFaceName(FileExt)
	Randomize
	RanNum = Int(90000*rnd)+10000
 	UserFaceName = UserID&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
End Function
Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function
%>
