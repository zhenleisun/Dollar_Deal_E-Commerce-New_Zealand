<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
Dim Conn,Db,StartTime,Timeset,SqlNum,ConnStr
Timeset=0
StartTime = Timer()
SqlNum=0
Db="Data/bbs_cnhww.asp"

Sub ConnectionDatabase
	Set conn=Server.CreateObject("ADODB.Connection")
	ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(Db)
	On Error Resume Next
	Conn.Open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ�����л����ݿ����Ӵ���"
		Response.End
	End If
End Sub
%>