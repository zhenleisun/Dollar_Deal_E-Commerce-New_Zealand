<!--#include file="Inc/SysConfig.Asp"-->
<!--#include file="Inc/page_Cls.Asp"-->
<%
Dim Action,PageUrl,SqlWhere,Caption,Content,strPageInfo,TextInfo
If Not YxBBs.Founduser Then YxBBs.Error("����û��ע����ߵ�½��")
If YxBBs.BoardID>0 Then
	YxBBs.CheckBoard()
	SqlWhere="BoardID="&YxBBs.BoardID&" And IsDel=False And "
Else
	SqlWhere="IsDel=False And "
End If


Action=Lcase(Request.QueryString("Action"))
Select Case Action
Case"mysay"
	Call YxBBs.Head("�ҵ�����")
	SqlWhere=SqlWhere&"Name='"&YxBBs.MyName&"'"
Case"myreply"
	Call YxBBs.Head("�ҵĻظ�")
	SqlWhere=SqlWhere&"Name<>'"&YxBBs.MyName&"' and (TopicID in (select ReplyTopicID from [YX_bbs"&YxBBs.TB&"] where name='"&YxBBs.MyName&"'))"
Case"mygood"
        Call YxBBs.Head("�ҵľ�������")
	SqlWhere=SqlWhere&"name='"&YxBBs.MyName&"' and IsGood"
Case"hot"
	Call YxBBs.Head("����������")
	SqlWhere =SqlWhere& " Hits>100"
Case"good"
		Call YxBBs.Head("��������")
		SqlWhere =SqlWhere& " IsGood"
Case"new"
		Call YxBBs.Head("��������")
		SqlWhere =SqlWhere& " DATEDIFF('d',[LastTime],'"&YxBBs.NowBbsTime&"')<1"
Case"48h"
		Call YxBBs.Head("48Сʱ������")
		SqlWhere =SqlWhere& " DATEDIFF('d',[LastTime],'"&YxBBs.NowBbsTime&"')<2"
	Case"week"
		Call YxBBs.Head("һ��������")
		SqlWhere =SqlWhere& " DATEDIFF('d',[LastTime],'"&YxBBs.NowBbsTime&"')<7"
	
	Case"reply"
		Call YxBBs.Head("�����ظ�����")
		SqlWhere =SqlWhere& " ReplyNum>9"
Case Else
	YxBBs.Error("����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������")
End Select
If YxBBs.BoardID=0 Then
	PageUrl="?Action="&Action
Else
 	PageUrl="?Action="&Action&"&BoardID="&YxBBs.BoardID
End If
Response.Write ShowTopic()&ShowListBottom()
YxBBs.Footer()
Set YxBBs=Nothing



Function ShowListBottom()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("�����б��ײ�")
	Temp=Replace(Temp,"{��ҳ}",strPageInfo)
	ShowListBottom=Temp
End Function

Function ShowTopic()
	Dim intPageNow,arr_Rs,i,pages,TopicImg,Conut,page
	Dim Temp,TopicS,Caption,Facepic,Moodpic,LastRe,RePageUrl,UploadType,RePage,leftn,ii
	Temp=YxBBs.Template.ReadTemplate("�����б�����")
	TopicS=Replace(Temp,"{����}",YxBBs.Stats )
	intPageNow = Request.QueryString("page")
	Set pages = New Cls_PageView
	pages.strTableName = "[YX_Topic]"
	pages.strPageUrl = PageUrl
	pages.strFieldsList = "Top 200 TopicID,Face,Caption,Name,TopType,IsGood,AddTime,BoardID,LastTime,Hits,LastReply,UploadType,IsVote,ReplyNum,SqlTableID,IsLock"
	pages.strCondiction = SqlWhere
	pages.strOrderList = "TopType desc,LastTime desc"
	pages.strPrimaryKey = "TopicID"
	pages.intPageSize = Int(YxBBs.BBSSetting(26))
	pages.intPageNow = intPageNow
	pages.strCookiesName = "BBSList"&Action&YxBBs.BoardID
	pages.Reloadtime=3
	pages.strPageVar = "page"
	'pages.intRefresh = 0
	pages.InitClass
	Arr_Rs = pages.arrRecordInfo
	strPageInfo = pages.strPageInfo
	Set pages = nothing
	If IsArray(Arr_Rs) Then
	TopicImg=Split(YxBBs.Template.ReadTemplate("����״̬ͼƬ"),"|")
	For i = 0 to UBound(Arr_Rs, 2)
	Moodpic=TopicImg(6)
	If Arr_Rs(13,i) > Int(YxBBs.BBSSetting(24)) Then Moodpic=TopicImg(5)
	If Arr_Rs(5,i) Then Moodpic=TopicImg(3)'����
	If Arr_Rs(15,i) Then Moodpic=TopicImg(7)'����
	If Arr_Rs(12,i) Then Moodpic=TopicImg(4)'ͶƱ
	If Arr_Rs(4,i)=5 Then Moodpic=TopicImg(0)'�ܶ�
	If Arr_Rs(4,i)=4 Then Moodpic=TopicImg(1)
	If Arr_Rs(4,i)=3 Then Moodpic=TopicImg(2)'��
	                Facepic=""
                                If Arr_Rs(1,i)<>99 then Facepic="<img src='Images/face/"&Arr_Rs(1,i)&".gIf'>"
				UploadType=""
				If Arr_Rs(11,i)<>"" Then Facepic="<img title='�����Ӻ��и���' src='Inc/Editor/Images/common.gif' border=0> "
	LastRe=split(Arr_Rs(10,i),"|")
	RePage=(Arr_Rs(13,i)+1)\10
	If RePage<(Arr_Rs(13,i)+1)/10 Then RePage=RePage+1
	RePageUrl="Show.Asp?ID="&Arr_Rs(0,i)
	leftn=50
	If RePage>1 Then leftn=36
	If Repage>10 Then leftn=30
	Caption=UploadType
	Caption=Caption&"<a href="&Repageurl&" title='���ⷢ��ʱ�䣺"&Arr_Rs(6,i)&"<br>���ظ����ݣ�"&YxBBs.Fun.StrLeft(LastRe(1),20)&"'>"&YxBBs.Fun.StrLeft(Arr_Rs(2,i),leftn)&"</a>"
	If Repage>1 Then
		Caption=Caption&" <a href='"&RePageurl&"&page="&RePage&"'>[<font color=#FF0000>"&RePage&"</font>]</a>"
	End If
	If Datediff("n",Arr_Rs(8,i),YxBBs.NowbbsTime)<=180 Then Caption=Caption&TopicImg(8)
	Temp=YxBBs.Template.ReadTemplate("�����б�")
	Temp=Replace(Temp,"{״̬}",Moodpic)
	Temp=Replace(Temp,"{����}",Caption)
	Temp=Replace(Temp,"{����}",Facepic)
	Temp=Replace(Temp,"{�û�����}","<a href='Profile.Asp?name="&Arr_Rs(3,i)&"' title='�鿴 "&Arr_Rs(3,i)&" ������'>"&Arr_Rs(3,i)&"</a>")
	Temp=Replace(Temp,"{�ظ���}",Arr_Rs(13,i))
	Temp=Replace(Temp,"{�����}",Arr_Rs(9,i))
	Temp=Replace(Temp,"{���ʱ��}","<font title='ʱ�䣺"&FORMATDATETIME(Arr_Rs(8,i),3)&"'>"&Datepart("yyyy",Arr_Rs(8,i))&"-"&Right(00&Datepart("m",Arr_Rs(8,i)),2)&"-"&Right(00&Datepart("D",Arr_Rs(8,i)),2)&"</font>")
	Temp=Replace(Temp,"{�ظ��û�����}",LastRe(0))
	TopicS=TopicS&Temp
	Next
	End If
	ShowTopic=TopicS
End Function

%>