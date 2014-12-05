<!--#include file="Inc/SysConfig.Asp"-->
<!--#include file="Inc/ubb_Cls.Asp"-->
<!--#include file="Inc/page_Cls.Asp"-->
<script type="text/javascript">
var y=0;function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){y++;if (y>1) {alert('帖子正在发出，请耐心等待！');return false;}this.document.yimxu.submit();this.document.yimxu.BtnPost.disabled=true;}}
</script>
<%
Dim TopicCaption,TopicUserName,TopicTopType,TopicIsGood,TopicIsLock,TopicIsVote,TopicHits,TopicReplyNum
Dim YxBBs_ubb,Page,PageString,PageInfo,ID,UserName,BBSID,GCode
ID=YxBBs.CheckNum(Request.QueryString("ID"))
Call GetTopicInfo()
Call ShowListTop()
If TopicIsVote Then ShowVote()
Set YxBBs_ubb=New YxBBsubb_Cls
Call ShowBBS()
Set YxBBs_ubb=Nothing
ShowListBottom()
FastReply()
YxBBs.Footer()
YxBBs.Execute("UPDATE [YX_Topic] SET Hits=Hits+1 WHERE TopicID="&ID&"")
Set YxBBs=Nothing
	
Sub GetTopicInfo()
	Dim Rs,Arr_Rs
	Set Rs=YxBBs.Execute("Select TopicID,Caption,Name,TopType,IsGood,IsLock,isVote,Hits,ReplyNum,AddTime,SqlTableID,BoardID From [YX_Topic] where TopicID="&ID&" and IsDel=False")
	IF Rs.eof then YxBBs.Error("该帖子不存在或已经删除！")
	Arr_Rs=Rs.GetRows(1)
	Rs.Close
	Set Rs=Nothing
	TopicCaption=YxBBs.Fun.HtmlCode(Arr_Rs(1,0))
	TopicUserName=Arr_Rs(2,0)
	TopicTopType=Arr_Rs(3,0)
	TopicIsGood=Arr_Rs(4,0)
	TopicIsLock=Arr_Rs(5,0)
	TopicIsVote=Arr_Rs(6,0)
	TopicHits=Arr_Rs(7,0)
	TopicReplyNum=Arr_Rs(8,0)
	YxBBs.TB=Arr_Rs(10,0)
	YxBBs.BoardID=Arr_Rs(11,0)
	YxBBs.CheckBoard()
	YxBBs.Head(Arr_Rs(1,0))
	If YxBBs.MyName<>TopicUserName Then
		If Cint(YxBBs.ClassSetting(1))=0 Then YxBBs.Error("您没有查看别人帖子的权限！")
		If TopicIsGood Then If Cint(YxBBs.ClassSetting(2))=0 Then YxBBs.Error("您没有浏览精华帖子的权限！")
	End If
End Sub

Function TopicMood()
	Dim Temp,TopicImg
	TopicImg=Split(YxBBs.Template.ReadTemplate("帖子状态图片"),"|")
	IF TopicIsGood Then Temp="<td>"&TopicImg(3)&" <font color=red>精华帖子</font></td>"
	IF TopicIsVote then Temp=Temp&"<td>"&TopicImg(4)&" <font color=red>投票帖子</font></td>"
	IF TopicIsLock then Temp=Temp&"<td>"&TopicImg(7)&" <font color=red>此帖被锁</font></td>"
	IF TopicTopType=3 then Temp=Temp&"<td>"&TopicImg(2)&" <font color=red>置顶帖子</font></td>"
	IF TopicTopType=4 then Temp=Temp&"<td>"&TopicImg(1)&" <font color=red>区置顶帖子</font></td>"
	IF TopicTopType=5 then Temp=Temp&"<td>"&TopicImg(0)&" <font color=red>总置顶帖子</font></td>"
	IF Temp<>"" Then Temp="<Div  align='right'><table border=1 style='border-collapse: collapse' cellpadding='5' cellspacing='0' bordercolor=#DEDEDE bgcolor=ffffff><tr>"&Temp&"</tr></table></div>"
	TopicMood=Temp
End Function

Function ShowListTop()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("帖子顶部")
	Temp=Replace(Temp,"{发帖按钮}",SayBar())
	Temp=Replace(Temp,"{点击数}",TopicHits+1)
	Temp=Replace(Temp,"{标题}",TopicCaption)
	Response.Write(Temp)
End Function

Function ShowListBottom()
	Dim Temp
	Temp=YxBBs.Template.ReadTemplate("帖子底部")
	Temp=Replace(Temp,"{分页}",PageInfo)
	Response.Write(Temp)
End Function

Function SayBar()
	Dim Temp,PostPic
	PostPic=Split(YxBBs.Template.ReadTemplate("发帖图片"),"||")
	Temp="<a href='Say.Asp?BoardID="&YxBBs.BoardID&"'>"&PostPic(0)&"</a> <a href='Say.Asp?action=vote&BoardID="&YxBBs.BoardID&"'>"&PostPic(1)&"</a>"
	If Not TopicIsLock or YxBBs.ClassID=1 Then
		Temp=Temp&" <a href='Say.Asp?Action=Reply&BoardID="&YxBBs.BoardID&"&ID="&ID&"'>"&PostPic(2)&"</a>"
	End If
	SayBar=Temp
End Function

Function ShowVote()
	Dim Temp,Rs,Arr_Rs,Vote,VoteNum,AllvoteNum,VotePicW,Opt,ClueTxt,CanVote,ii,VoteShow,i
	Set Rs=YxBBs.Execute("Select TopicID,Vote,VoteNum,VoteType,OutTime From [YX_TopicVote] where TopicID="&ID&"")
	If Rs.Eof then Exit Function
	Arr_Rs=Rs.GetRows(1)
	Rs.Close:Set Rs=Nothing
	Vote=Split(Arr_Rs(1,0),"|")
	VoteNum=split(Arr_Rs(2,0),"|")
	CanVote=True
	If Not YxBBs.FoundUser Then
		ClueTxt="你还没有登陆，不能进行投票。"
		CanVote=False
	Else
		IF not YxBBs.Execute("Select User From [YX_TopicVoteUser] where User='"&YxBBs.MyName&"' and TopicID="&ID&"").eof then
			ClueTxt="你已经投过票了，不能再投票了。[<a style='cursor:hand;color:red;' onclick=javascript:window.open('See.Asp?Action=Vote&Id="&Arr_Rs(0,0)&"','open','width=550,height=400,resizable=1,scrollbars=yes,menubar=no,status=yes')>详情</a>]"
			CanVote=False
		End If
	End If
	For i=1 to ubound(Vote)
		AllvoteNum=Int(AllvoteNum+VoteNum(i))
	Next
	IF AllVoteNum=0 then AllvoteNum=1
	For i=1 To ubound(Vote)
		ii=ii+1
		VotePicW=VoteNum(i)/AllvoteNum*85
		If CanVote Then
			IF Int(Arr_Rs(3,0))=1 then Opt="<input type=radio value='"&i&"' name='opt'>" Else Opt="<input type=checkbox name='opt"&i&"' value='"&i&"'>"
		End If
		dim Svote
		IF not YxBBs.Execute("Select User From [YX_TopicVoteUser] where User='"&YxBBs.MyName&"' and TopicID="&ID&"").eof then
		Svote="<img border=0 height=8 width='"&VotePicW&"%' src='Images/hr.gif'> <b>"&VoteNum(i)&"</b> 票"
		else
                Svote="<img border=0 height=8 width='50%' src='Images/hr.gif'>  投票后可查看详情！"
		End If
		VoteShow=VoteShow&"<tr><td width='5%' height=25 >&nbsp;"&i&"</td><td width='40%'>&nbsp;"&Opt&""&YxBBs.Fun.HtmlCode(Vote(i))&"</td><td width='35%' >&nbsp;"&Svote&"</td></tr>"

	Next
	If DateDiff("s",YxBBs.NowBbsTime,Arr_Rs(4,0))<0 then
		ClueTxt="该投票已经过期，不能进行投票。"
		CanVote=False
	End If
	IF CanVote then
		ClueTxt="<input type='submit' value='投 票 (投票后方能看详尽结果)' ID='BtnPost'>"
	End IF
	ClueTxt=ClueTxt&" [ 截止时间："&Arr_Rs(4,0)&" ]"
	Temp=YxBBs.Template.ReadTemplate("显示投票")
	Temp=Replace(Temp,"{投票选项}","<form style='margin:0' method=POST onsubmit='BtnPost.disabled = true;' action=VoteSubmit.Asp?Action=Vote&id="&ID&"&type="&Arr_Rs(3,0)&">"&VoteShow)
	Temp=Replace(Temp,"{投票信息}",ClueTxt&"</form>")
	Response.Write(Temp)
End Function

Function ShowBBS()
	Dim arr_Rs,i,pages,rs,GradeName,Gradepic
	Dim Template,TempUserL,Temp,TempStr,Lou,Sign,ShowCaption,ShowMood,Arr_Ip
	Page = Request.QueryString("page")
	Set pages = New Cls_PageView
	pages.strTableName = "[YX_Bbs"&YxBBs.TB&"] As B inner join [YX_User] As U on B.Name=U.Name"
	pages.strPageUrl = "?BoardID="&YxBBs.BoardID&"&ID="&ID&"&TB="&YxBBs.TB
	pages.strFieldsList = "B.BbsID,B.TopicID,B.Caption,B.Content,B.Name,B.AddTime,B.BoardID,B.UbbString,B.IP,U.Id,U.Name,U.IsQQpic,U.QQ,U.Pic,U.Picw,U.Pich,U.GradeNum,U.EssayNum,U.Mail,U.Home,U.Sex,U.Mark,U.Coin,U.Sign,U.Regtime,U.IsShow,U.IsSign,U.RegIp,U.LoginNum,U.Honor,U.TimeSum"
	pages.strCondiction = "B.isDel=False and (B.TopicID="&ID&" or B.ReplyTopicID="&ID&")"
	pages.strOrderList = "B.BbsID"
	pages.strPrimaryKey = "BbsID"
	Pages.CountSQL=TopicReplyNum+1
	'Pages.CountSQL=YxBBs.Execute("SELECT Count(ID) FROM [YX_Bbs"&YxBBs.TB&"] where TopicID="&ID&" or ReplyTopicID="&ID&"")(0)
	pages.intPageSize = 10
	pages.intPageNow = Page
	pages.strCookiesName = "Show_"&ID
	pages.Reloadtime=0
	pages.strPageVar = "page"
	pages.InitClass
	Arr_Rs = pages.arrRecordInfo
	PageInfo = pages.strPageInfo
	page=pages.intPageNow
	Set pages = nothing
	Template=YxBBs.Template.ReadTemplate("帖子信息")
	
	If IsArray(Arr_Rs) Then
		Lou = (Page-1)*10
		For i = 0 to UBound(Arr_Rs,2)
			UserName=Arr_Rs(10,i)
			BBSID=Arr_Rs(0,i)
			YxBBs_ubb.UbbString=Arr_Rs(7,i)
			Temp=Replace(Template,"{信息}","")
			Temp=Replace(Temp,"{短信}","")
			Temp=Replace(Temp,"{QQ}",YxBBs.Fun.GetSqlStr(Arr_Rs(12,i)))
			Temp=Replace(Temp,"{邮箱}",YxBBs.Fun.GetSqlStr(Arr_Rs(18,i)))
			Temp=Replace(Temp,"{主页}",YxBBs.Fun.GetSqlStr(Arr_Rs(19,i)))
			Temp=Replace(Temp,"{编辑}","")
			Temp=Replace(Temp,"{引用}","")
			Temp=Replace(Temp,"{楼号}","<b><font color=#FF0000>"&Lou+i+1&"</font><sup>F</sup></b>")
			dim ad
			SET ad=YxBBs.Execute("Select google From [YX_config]")
			If Arr_Rs(1,i)=ID Then
				Temp=Replace(Temp,"{删除}",SetTopic())
				Temp=Replace(Temp,"{广告}",ad(0))
				ShowCaption=TopicCaption
				ShowMood=TopicMood()
			Else
				ShowCaption="":ShowMood=""
				Temp=Replace(Temp,"{广告}","")
				Temp=Replace(Temp,"{删除}","<a title=删除此回复 href=admin.Asp?BoardID={版块ID}&ID={主题ID}&BbsID={帖子ID}&TB={数据表ID}&Action=删除回复&>×</a>")
			End If
			Temp=Replace(Temp,"{数据表ID}",YxBBs.TB)
			Temp=Replace(Temp,"{版块ID}",YxBBs.BoardID)
			Temp=Replace(Temp,"{主题ID}",ID)
			Temp=Replace(Temp,"{页次}",page)
			Temp=Replace(Temp,"{帖子ID}",Arr_Rs(0,i))
			Temp=Replace(Temp,"{用户名称}",Arr_Rs(10,i))
			Temp=Replace(Temp,"{发帖时间}",Arr_Rs(5,i))
			SET Rs=YxBBs.Execute("Select GradeName,GradePic From [YX_UserGrade] where GradeNum="&Arr_Rs(16,i)&"")
            Temp=Replace(Temp,"{等级图片}","<img src='"&Rs("Gradepic")&"' alt="&Rs("GradeName")&" Align=center>")
			Temp=Replace(Temp,"{等级名称}",Rs("GradeName"))
			IF YxBBs.ClassID>2 then
				Temp=Replace(Temp,"{用户IP}",YxBBs.Fun.GetIp(Arr_Rs(8,i),2))
			Else
				Temp=Replace(Temp,"{用户IP}",YxBBs.Fun.GetSqlStr(Arr_Rs(8,i)))
			End If
			Temp=Replace(Temp,"{Y币}",""&YxBBs.BBSSetting(23)&"：<b><font color=red>"&YxBBs.Fun.GetSqlStr(Arr_Rs(21,i))&"</font></b>")
			Temp=Replace(Temp,"{金钱数}",YxBBs.Fun.GetSqlStr(Arr_Rs(22,i)))
			
			If Arr_Rs(29,i)<>"" Then
				Temp=Replace(Temp,"{头衔}","头衔："&YxBBs.Fun.GetSqlStr(Arr_Rs(29,i))&"<br>")
			Else
				Temp=Replace(Temp,"{头衔}","头衔：未定义<br>")
			End If
			If IsOnline(Arr_Rs(10,i)) Then
				If Arr_Rs(20,i) Then
					Temp=Replace(Temp,"{性别}","<img src='Images/11.gif' alt='性别：男<br>状态：在线<br>积时：" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				Else
					Temp=Replace(Temp,"{性别}","<img src='Images/01.gif' alt='性别：女<br>状态：在线<br>积时：" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				End If
			Else
				If Arr_Rs(20,i) Then
					Temp=Replace(Temp,"{性别}","<img src='Images/10.gif' alt='性别：男<br>状态：离线<br>积时：" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				Else
					Temp=Replace(Temp,"{性别}","<img src='Images/00.gif' alt='性别：女<br>状态：离线<br>积时：" &YxBBs.Fun.DateToSting(Arr_Rs(30,i))&"' align='absmiddle'>")
				End If
			End If
			Temp=Replace(Temp,"{帖数}",Arr_Rs(17,i))
			Temp=Replace(Temp,"{注册时间}",formatdatetime(Arr_Rs(24,i),2))
			IF Arr_Rs(11,i) then
				Temp=Replace(Temp,"{用户头像}","<img src='http://qqshow-user.tencent.com/"&Arr_Rs(12,i)&"/10/'>")
			else
				Temp=Replace(Temp,"{用户头像}","<img border='0' src='"&Arr_Rs(13,i)&"' width="&Arr_Rs(14,i)&" height="&Arr_Rs(15,i)&">")
			End If
			Sign=Arr_Rs(23,i)
			IF Arr_Rs(25,i) then
				Temp=Replace(Temp,"{帖子内容}","<Hr Size=1><Font Color=red>该用户帖子已被管理员屏蔽，请和管理员联系</Font><Hr Size=1>")
				Temp=Replace(Temp,"{签名}","【签名档被已管理员屏蔽】")
			Else
				IF YxBBs.BBSSetting(12)="1" then Sign="" 
				IF Arr_Rs(26,i) Then Sign="【签名档已被管理员屏蔽】"
				IF isNull(sign) or sign="" then
					Temp=Replace(Temp,"{签名}",YxBBs.BBSName&"欢迎您！")
				Else
					Temp=Replace(Temp,"{签名}",YxBBs_ubb.Sign_Code(Sign))
				End IF
				Temp=Replace(Temp,"{帖子内容}",ShowMood&"<SPAN style='line-height:150%;table-layout:fixed;word-wrap:break-word;word-break:break-all;width:100%'><b>"&ShowCaption&"</b><br><br>"&YxBBs_ubb.UBB(Arr_Rs(3,i),1)&"</SPAN>")
			End IF
			TempStr=TempStr&Temp
		Next
		Response.Write(TempStr)
	End If
End Function

Function IsOnline(TUserName)
	Dim Temp,Rs,Ary_Rs,i
	IsOnline=false
 		Cache.Name="UserOnline"
	If Cache.valid Then
	   Temp=Cache.Value
	Else
		Set Rs=YxBBs.Execute("Select Name From [YX_online] where ClassID<>6")
		If Rs.Eof Or Rs.Bof Then Exit Function
		Ary_Rs=Rs.GetRows(-1)
		Rs.Close
		For i=0 To Ubound(Ary_Rs,2)
			Temp=Temp & Ary_Rs(0,i)&","
		Next
		Temp=","&Temp
		Cache.add Temp,dateadd("n",5,YxBBs.NowBBSTime)'5分钟
	End If
	If InStr(Lcase(Temp),","&Lcase(TUserName)&",")>0 Then IsOnline=true
End Function

Function SetTopic()
	Dim Temp,Url
	If Not YxBBs.FoundUser Then Exit Function
	Url="admin.Asp?BoardID="&YxBBs.BoardID&"&ID="&ID&"&TB="&YxBBs.TB&"&Action="
        
	IF TopicTopType=5 then Temp="<div class=menuitems><a href="&Url&"取消总置顶>取消总置顶</a></div>" Else Temp="<div class=menuitems><a href="&Url&"总置顶>设为总置顶</a></div>"
	IF TopicTopType=4 Then Temp=Temp&"<div class=menuitems><a href="&Url&"取消区置顶>取消区置顶</a></div>" ELse Temp=Temp&"<div class=menuitems><a href="&Url&"区置顶>设为区置顶</a></div>"
	IF TopicTopType=3 Then Temp=Temp&"<div class=menuitems><a href="&Url&"取消置顶>取消置顶</a></div>" ELse Temp=Temp&"<div class=menuitems><a href="&Url&"置顶>设为置顶</a></div>"
	IF TopicIsGood Then Temp=Temp&"<div class=menuitems><a href="&Url&"取消精华>取消精华</a></div>" Else Temp=Temp&"<div class=menuitems><a href="&Url&"精华>设为精华</a></div>"
	IF TopicIsLock Then Temp=Temp&"<div class=menuitems><a href="&Url&"解锁>解除锁定</a></div>" Else Temp=Temp&"<div class=menuitems><a href="&Url&"锁定>锁定主题</a></div>"
	Temp=Temp&"<div class=menuitems><a href="&Url&"删除主题>删除主题</a></div><div class=menuitems><a href="&Url&"移动>移动主题</a></div><div class=menuitems><a href="&Url&"提升>提升主题</a></div>"
       Temp="<a onmouseover=""showmenu(event,'"&Temp&"')"">管理</a>"
	SetTopic=Temp
End Function

Function FastReply()
	If Not YxBBs.FoundUser Then Exit Function
	IF Not TopicIsLock Or YxBBs.ClassID=1 then
		Dim Caption,Content
		Caption="快速回复【<a href='Say.Asp?action=reply&BoardID="&YxBBs.BoardID&"&id="&ID&"'>高级回复</a>】"
		Content="<form style='margin: 8;' method='POST' name='yimxu' onsubmit='if(content.value==""""){alert(""您还没有填写帖子内容!"");content.focus();return (false);};"
		If Cint(YxBBs.BoardSetting(0))=1 then Content=Content&"if(CheckCode.value==""""){alert(""请填写验证码！"");CheckCode.focus();return (false);};"
		Content=Content&"BtnPost.disabled = true;' action='Save.Asp?action=Reply&BoardID="&YxBBs.BoardID&"&TB="&YxBBs.TB&"&ID="&ID&"&page="&Page&"'><table align='center' border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='100%'><tr><td width='15%' valign='top'><p style='line-height: 150%; margin: 6'>回复说明：<br>论坛是我家,维护靠大家,请遵守国家法规,不乱发水帖!</p>"&_
		"<td width='76%' valign='top'><input type=hidden name='caption' value='Re:"&TopicCaption&"'><textarea onkeydown=presskey(); rows=6 name=content Style='width:98%'></textarea><p style='margin-top: 10; margin-bottom: 0'>"
		IF Cint(YxBBs.BoardSetting(0))=1 then Content=Content&""&YxBBs.GCode&""
		Content=Content&"<input type=submit value='回复帖子' ID='BtnPost'> <input type='reset' value='重 写'> 快捷键 Ctrl+Enter </p></td><td width='5%' valign='top'></td></tr></table></form>"
	Call YxBBs.ShowTable(Caption,Content)
	END IF
End Function
%>