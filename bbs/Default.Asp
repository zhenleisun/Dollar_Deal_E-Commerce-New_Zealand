<!--#include file="Inc/SysConfig.Asp"-->

<%
Dim PageString,GCode
YxBBs.Head("��̳��ҳ")
If YxBBs.BBSSetting(7)="0" then PageString=ShowInfo()
PageString=PageString&ShowBoard()
If YxBBs.BBSSetting(10)="0" then PageString=PageString&ShowOnline()
Response.Write PageString &"<iframe name='hiddenframe' frameborder='0' height='0' id='hiddenframe'></iframe>"
If YxBBs.BBSSetting(32)="0" then ShowBirthday()
If YxBBs.BBSSetting(8)="0" then ShowLink()
YxBBs.Footer()
Set Cache = Nothing
Set YxBBs = Nothing

Function ShowInfo()
	Dim Temp,OnlingType
	If YxBBs.FoundUser Then
		Temp = YxBBs.Template.ReadTemplate("�û���Ϣ")
		If YxBBs.MyIsQQpic Then
			Temp=Replace(Temp,"{�û�ͷ��}","<img src='http://qqshow-user.tencent.com/"&YxBBs.MyQQ&"/11/'>")
		Else
			Temp=Replace(Temp,"{�û�ͷ��}","<img src="&YxBBs.Mypic&" width="&YxBBs.Mypicw&" height="&YxBBs.Mypich&" >")
		End if
		OnlingType="<font Color=Red>�����ߣ�</Font>"
		If YxBBs.MyHidden=2 Then OnlingType="<font Color=Red>��������</Font>"
		Temp=Replace(Temp,"{�û�����}","<a href='Profile.Asp' title='���Լ���������Ϣ'>"&YxBBs.MyName&"</a>")
		Temp=Replace(Temp,"{����״̬}",OnlingType)
		Temp=Replace(Temp,"{��Ǯ��}",YxBBs.MyCoin)
		Temp=Replace(Temp,"{Y��}",""&YxBBs.BBSSetting(23)&"��<font color=red>"&YxBBs.MyMark&"</font>")
		Temp=Replace(Temp,"{����}",YxBBs.MyEssayNum)
		Temp=Replace(Temp,"{�ȼ�����}",YxBBs.MyGradeName)
	Else
		Temp = YxBBs.Template.ReadTemplate("���ٵ�½")
	End If 
	Temp=Replace(Temp,"{����}",YxBBs.Srule)
        Temp=Replace(Temp,"{�»�Ա����}",YxBBs.NewUser) 
	Temp=Replace(Temp,"{��������}",YxBBs.TodayNum)
	Temp=Replace(Temp,"{���������}",YxBBs.MaxEssayNum)
	Temp=Replace(Temp,"{������}",YxBBs.TopicNum)
	Temp=Replace(Temp,"{������}",YxBBs.AllEssayNum)
	Temp=Replace(Temp,"{��Ա��}",YxBBs.UserNum)
	Temp=Replace(Temp,"{��������}",YxBBs.YsterdayNum)
        If Cint(YxBBs.BBSSetting(6))=0 Then 
        Temp=Replace(Temp,"{��֤��}",YxBBs.GCode)
        else
        Temp=Replace(Temp,"{��֤��}","")
        end if
	ShowInfo=Temp
End Function

Function ShowBoard()
	Dim Rs,Temp,i,BigBoard,BoardChild,BoardBottom,BoardStr
	BigBoard=YxBBs.Template.ReadTemplate("��������")
	BoardChild=YxBBs.Template.ReadTemplate("����б�")
	BoardBottom=YxBBs.Template.ReadTemplate("�����ײ�����")
	If Not IsArray(YxBBs.Board_Rs) Then YxBBs.CacheBoard()
	If Not IsArray(YxBBs.Board_Rs) Then Exit Function
	For i=0 To Ubound(YxBBs.Board_Rs,2)
		If YxBBs.Board_Rs(0,i)<2 Then
			Temp=""
			If YxBBs.Board_Rs(0,i)=0 Then
				If i >= 1 Then Temp = Temp & BoardBottom & VbCrlf
				Temp = Temp & Replace(BigBoard,"{��������}",YxBBs.Board_Rs(3,i))
			Else
				Temp=BoardChild
				Temp=YxBBs.GetBoardInfo(Temp,i)
			End If
			BoardStr=BoardStr & Temp
		End If
	Next
	ShowBoard = BoardStr& BoardBottom & VbCrlf
End Function

Function ShowOnline()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("����ͳ��")
	Temp=Replace(Temp,"{�����б�}","<span id=showon></span>")
	Dim BrowserType
	Set BrowserType=New Cls_Browser
		Temp=Replace(Temp,"{�û�ϵͳ}",BrowserType.platform)
		Temp=Replace(Temp,"{�û������}",BrowserType.Browser & BrowserType.version)
	Set BrowserType=Nothing
	Temp=Replace(Temp,"{�û�IP}",YxBBs.MyIp)
	Temp=Replace(Temp,"{��վʱ��}",FORMATDATETIME(YxBBs.BuildDate,1))
	Temp=Replace(Temp,"{��������}",YxBBs.AllOnlineNum)
	Temp=Replace(Temp,"{��Ա��}",YxBBs.UserOnlineNum)
	Temp=Replace(Temp,"{�ο���}",YxBBs.AllOnlineNum-YxBBs.UserOnlineNum)
	Temp=Replace(Temp,"{���������}",YxBBs.MaxOnlineNum)
	Temp=Replace(Temp,"{�����ʱ��}",YxBBs.MaxOnlineTime)
	If Session(YxBBs.CacheName&"online")="1" then
		Temp=Replace(Temp,"{�����б�����}","<a target='hiddenframe' href='Online.Asp?Action=LoadIng' onClick='ShowOnline();' ><span id='Showtxt' >�ر���ϸ�б�</span></a>")
		Temp=Temp&"<iframe frameborder='0'  height='0' src='Online.Asp?Action=LoadIng&id=1'></iframe>"
	Else
		Temp=Replace(Temp,"{�����б�����}","<a target='hiddenframe' href='Online.Asp?Action=LoadIng' onClick='ShowOnline();' ><span id='Showtxt' >��ʾ��ϸ�б�</span></a>")
	End If
	ShowOnline=Temp
End Function

Sub ShowLink()
	Dim Link_List,LogoLink,TxtLink,Rs,Arr_Rs,I,Sql,j,k
	Cache.Name="Link_List"
	If Cache.Valid then
		Link_List=Cache.Value
	Else
		Sql = "Select BbsName,Url,Pic,Readme,IsPic From [YX_Link] Where  Orders  order by Orders Asc"
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Rs = Server.CreateObject("Adodb.RecordSet")
		Rs.open Sql,Conn,1,1
		SqlNum=SqlNum+1
		If Not(rs.eof or rs.bof) Then Arr_Rs = Rs.getrows()
		Rs.Close:Set Rs=Nothing
		If IsArray(Arr_Rs) Then
			TxtLink="":LogoLink="":j=0:k=0
			For i = 0 to UBound(Arr_Rs,2)
				If Arr_Rs(4,i) And Arr_Rs(2,i)<>"" Then
					j=j+1
					LogoLink=LogoLink&"<td align=""center""><a href="""&Arr_Rs(1,i)&""" target=""_blank""><img src="""&Arr_Rs(2,i)&""" border=""0"" alt=""::"&Arr_Rs(0,i)&"::"&Arr_Rs(3,i)&""" width=""88"" height=""31""></a></td>"
					If j=7 Then j=0:LogoLink=LogoLink&"</tr><tr>"
				Else
					k=k+1
					TxtLink=TxtLink&"<td align=""center""><a href="""&Arr_Rs(1,i)&""" target=""_blank"" title="""&Arr_Rs(3,i)&""">"&Arr_Rs(0,i)&"</a></td>"
					If k=7 Then k=0:TxtLink=TxtLink&"</tr><tr>"
				End If
			Next
		End If
		Link_List="<tr><td align=""center""><table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""0""><tr>"&TxtLink&"</tr></table><table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""0""><tr>"&LogoLink&"</tr></table></td></tr>"
		Cache.add Link_List,dateadd("n",5000,YxBBs.NowBBSTime)
	End If
	Call YxBBs.ShowTable("��̳����",Link_List)
End Sub

Sub ShowBirthday()
	Dim UserBirthday,temp,rs,Arr_Rs,i,Num
	Cache.Name="Birthday"
	If Cache.valid then
		Temp=Split(Cache.Value,"|")
		Num=Temp(0)
		UserBirthday=Temp(1)
	Else
		Set Rs=YxBBs.Execute("Select Name,Birthday From[YX_User] where Month(Birthday)=Month(now) and day(Birthday)=day(now) and Birthday<>Cdate('1900-1-1')")
		IF Not Rs.eof Then Arr_Rs=Rs.getrows()
		Rs.Close
		Num=0
		If IsArray(Arr_Rs) Then
			For i = 0 to UBound(Arr_Rs,2)
				Num=Num+1
				UserBirthday=UserBirthday&"  <a href=Profile.Asp?Name="&Arr_Rs(0,i)&"><img src=Images/Birthday.gif align=""absmiddle"" border=0> ף��"&Arr_Rs(0,i)&"�����տ���</a>  "
			Next
		End If
		If i>3 Then UserBirthday="<marquee onMouseOver='this.stop()' onMouseOut='this.start()' scrollamount='3' >"&UserBirthday&"</marquee>"
		Temp=Num&"|"&UserBirthday
		Cache.add Temp,dateadd("n",100,now)
	End If
	If num > 0 Then
		Call YxBBs.ShowTable("���칲�� "&num&" λ��Ա������","<tr><td height=30>"&UserBirthday&"</td></tr>")
	End If
End Sub
%>
<%
Class Cls_Browser
	Public Browser,version,platform,IsSearch
	Private Sub Class_Initialize()
		Dim Agent,Tmpstr
		Agent=Request.ServerVariables("HTTP_USER_AGENT")
		'Agent="Opera/7.23 (X11; Linux i686; U)  [en]"
		If Left(Agent,7) ="Mozilla" Then
			Agent=Split(Agent,";")
			If InStr(Agent(1),"MSIE")>0 Then
				Browser="Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
                        If InStr(Agent(2),"NT 6.0")>0 Then 
				platform="Windows vista"
			ElseIf InStr(Agent(2),"NT 5.2")>0 Then
				platform="Windows 2003"
			ElseIf InStr(Agent(2),"Windows CE")>0 Then
				platform="Windows CE"
			ElseIf InStr(Agent(2),"NT 5.1")>0 Then
				platform="Windows XP"
			ElseIf InStr(Agent(2),"NT 4.0")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(2),"NT 5.0")>0 Then
				platform="Windows 2000"
			ElseIf InStr(Agent(2),"NT")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(2),"9x")>0 Then
				platform="Windows ME"
			ElseIf InStr(Agent(2),"98")>0 Then
				platform="Windows 98"
			ElseIf InStr(Agent(2),"95")>0 Then
				platform="Windows 95"
			ElseIf InStr(Agent(2),"Win32")>0 Then
				platform="Win32"
			ElseIf InStr(Agent(2),"Linux")>0 Then
				platform="Linux"
			ElseIf InStr(Agent(2),"SunOS")>0 Then
				platform="SunOS"
			ElseIf InStr(Agent(2),"Mac")>0 Then
				platform="Mac"
			ElseIf UBound(Agent)>2 Then
				If InStr(Agent(3),"NT 5.1")>0 Then
					platform="Windows XP"
				End If 
				If InStr(Agent(3),"Linux")>0 Then
					platform="Linux"
				End If
			End If
			If InStr(Agent(2),"Windows")>0 And platform="unknown" Then
				platform="Windows"
			End If
		ElseIf Left(Agent,5) ="Opera" Then
			Agent=Split(Agent,"/")
			Browser="Mozilla "
			tmpstr=Split(Agent(1)," ")
			version=tmpstr(0)
                        If InStr(Agent(1),"NT 6.0")>0 Then 
				platform="Windows vista"
			ElseIf InStr(Agent(1),"NT 5.2")>0 Then
				platform="Windows 2003"
			ElseIf InStr(Agent(1),"Windows CE")>0 Then
				platform="Windows CE"
			ElseIf InStr(Agent(1),"NT 5.1")>0 Then
				platform="Windows XP"
			ElseIf InStr(Agent(1),"NT 4.0")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(1),"NT 5.0")>0 Then
				platform="Windows 2000"
			ElseIf InStr(Agent(1),"NT")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(1),"9x")>0 Then
				platform="Windows ME"
			ElseIf InStr(Agent(1),"98")>0 Then
				platform="Windows 98"
			ElseIf InStr(Agent(1),"95")>0 Then
				platform="Windows 95"
			ElseIf InStr(Agent(1),"Win32")>0 Then
				platform="Win32"
			ElseIf InStr(Agent(1),"Linux")>0 Then
				platform="Linux"
			ElseIf InStr(Agent(1),"SunOS")>0 Then
				platform="SunOS"
			ElseIf InStr(Agent(1),"Mac")>0 Then
				platform="Mac"
			ElseIf UBound(Agent)>2 Then
				If InStr(Agent(3),"NT 5.1")>0 Then
					platform="Windows XP"
				End If 
				If InStr(Agent(3),"Linux")>0 Then
					platform="Linux"
				End If
			End If
		Else
			'ʶ����������
			Dim botlist,i
			Botlist="Google,Isaac,SurveyBot,Baiduspider,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir"
			Botlist=split(Botlist,",")
			For i=0 to UBound(Botlist)
				If InStr(Agent,Botlist(i))>0  Then 
					platform=Botlist(i)&"������"
					IsSearch=True
					Exit For
				End If
			Next 
		End If
	End Sub
End Class
%>