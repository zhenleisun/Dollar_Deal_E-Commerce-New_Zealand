<!--#include file="Inc/SysConfig.Asp"-->
<%if Request("action")="" then
Dim TFilePath
On Error Resume Next
Response.Buffer = True
Response.Clear
Const FilePath = "UploadFile/TopicFile"
Const FacePath = "UploadFile/Head"

Sub ForbidMake()
	Dim Come,Here
	Come=Cstr(Request.ServerVariables("HTTP_REFERER"))
	Here=Cstr(Request.ServerVariables("SERVER_NAME"))
	If Come<>"" And Mid(Come,8,Len(Here)) <> Here Then
		Response.Redirect "Images/NoImg.gif"
		Response.end
	End If		
End Sub

Function ChkFile(FileName)
	Dim Temp,FileType,F
	ChkFile=false
	FileType=Lcase(Split(FileName,".")(ubound(Split(FileName,"."))))
	Temp="|asp|aspx|cgi|php|cdx|cer|asa|"
	If Instr(Temp,"|"&FileType&"|")>0 Then ChkFile=True
	F=Replace(Request("FileName"),".","")
	If instr(1,F,chr(39))>0 or instr(1,F,chr(34))>0 or instr(1,F,chr(59))>0 then ChkFile=True
End Function

Function DownloadFile(FileName)
Dim TempFileName,Fsize
    On error resume next
	Server.ScriptTimeOut=999999
	Response.Clear
    Dim FileType,ADS,StrFileName,Data
    FileType=Lcase(Split(FileName,".")(ubound(Split(FileName,"."))))
	StrFileName=Server.Mappath(FileName)
	TempFileName = Split(StrFileName,"\")(Ubound(Split(StrFileName,"\")))
    Set ADS = Server.CreateObject("ADODB.Stream") 
	ADS.Open
	ADS.Type = 1 
    ADS.LoadFromFile(StrFileName)
	Data=ADS.Read
	Fsize=Clng(lenb(Data))
	If Err Then
  	   Response.Redirect("Images/NoImg.gif")
	   'Response.Write("<h1>����: </h1>" & err.Description & "<p>")
       Response.End 
    End If
	ADS.Close
    If Response.IsClientConnected Then 
       If FileType="gif" Or FileType="jpg" Or FileType="jpeg" Or FileType="bmp" Then 
	Response.AddHeader "Content-Disposition", "attachment; filename=" & TempFileName
	      Response.ContentType = "image/*"
	   Else
	      Response.AddHeader "Content-Disposition", "attachment; filename=" & TempFileName
		  Response.ContentType = "application/ms-download"
	   End If
	   Response.AddHeader "Content-Length", Fsize
 	   Response.CharSet = "UTF-8" 
	   Response.ContentType = "application/octet-stream" 
	   Response.BinaryWrite Data
	   Response.Flush
	End If
End Function

If YxBBs.BBSSetting(61)="0" Then ForbidMake
If LCase(Request("Path"))="face" Then
   TFilePath=  FacePath & "/" & Request("FileName")
Else
   TFilePath = FilePath & "/" & Request("FileName")
End If

If ChkFile(TFilePath) Then
	Response.Redirect("Images/NoImg.gif")
End If
DownloadFile(TFilePath)
Set Yxbbs=Nothing
end if

if Request("action")="Upload" then%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><style>
BODY {
	FONT-SIZE: 12px;
	FONT-FAMILY: "����","Verdana", "Arial", "Helvetica", "sans-serif";
	margin-left: 0px;
}
Input {
	font-family: Arial;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	background-color: #F6F6F6;
	height: 18px;
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #999999;
	border-right-color: #EEEEEE;
	border-bottom-color: #EEEEEE;
	border-left-color: #999999;
}
</style></head><%
Dim Flag,BoardID
BoardID=Request.QueryString("BoardID")
If Request("Flag") = "" Then  Flag= 0 Else Flag = 1
Response.Write"<body topmargin='0' leftmargin='0'><form action='SaveUpload.Asp?BoardID="&BoardID&"&Flag="&Flag&"' method='post' enctype='multipart/form-data' target='_self'><input name='filedata' type='file' size='18'> <input type='Submit' value='�� ��' name='Submit' id='Submit'></form>"
%>
</body>
</html>
<%end if%>