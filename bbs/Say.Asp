<!--#include file="Inc/SysConfig.Asp"-->
<script type="text/javascript" src="inc/Editor/htmlEditor.js"></script>
<script type="text/javascript">
var editor = new htmlEditor("editor");
editor.hiddenName = "content";
editor.width = "100%";
editor.height = "250px";
editor.imagePath="inc/editor/images/";
editor.iconPath = 'inc/editor/images/face/';

</script><%
Dim Caption,Action,Submiturl,bbsCaption,ThisFace,IsTitle,BbsContent,Quote,Page,bbsVote,ID,Rs,rs1
If Not YxBBs.Founduser then YxBBs.Error("����û��ע����ߵ�½��")
YxBBs.CheckBoard()
ID=YxBBs.CheckNum(Request.QueryString("ID"))
Page=YxBBs.CheckNum(Request.QueryString("page"))
IsTitle=True
Action=lcase(Request.QueryString("action"))
Select Case Action
	Case"vote"
		Vote()
	Case"reply"
		Reply()
	Case"edit"
		Edit()
	Case Else
		YxBBs.Stats="��������"
		Caption=YxBBs.Stats
		bbsCaption="<input name='Caption' type='text' style='width:80%' maxlength=200>"
		Submiturl="Save.Asp?BoardID="&YxBBs.BoardID&""
End Select
YxBBs.Head(YxBBs.Stats)
Set Rs=YxBBs.Execute("Select Name,GradeNum From [YX_user]where Name='"&YxBBs.MyName&"' And PassWord='"&YxBBs.MyPwd&"' And ID="&YxBBs.MyID&"")
ShowMain()
YxBBs.Footer()
Set YxBBs=Nothing

Sub Vote()
	Dim Temp,i
	Temp=YxBBs.BoardSetting(4)
	If Cint(YxBBs.ClassSetting(3))=0 Then YxBBs.Error("��û�в鿴ͶƱ��Ȩ�ޣ�")
	bbsVote="<tr><td valign=top><br><b>&nbsp;ͶƱѡ�<b></td><td>&nbsp;��ѡ��ͶƱ��Ŀ����<SELECT align=absmiddle name=VoteNum id=VoteNum onchange=SetNum(this)>"
	for i = 2 to int(Temp)
		bbsVote=bbsVote&"<option value='"&I&"'>"&I&"</option>"
	Next
	bbsVote=bbsVote&"</SELECT> ������ѡ <input type='checkbox' name='VoteType' value='2'> ����ʱ�䣺<select size='1' name='OutTime' ><option value='1'>һ��</option><option value='3'>����</option><option value='7'>һ��</option><option value='15'>�����</option><option value='31'>һ����</option><option value='93'>������</option><option value='365'>һ��</option><option value='10000' selected>������</option></select><hr size=1 width='98%'><span id=optionid><table cellpadding=0 cellspacing=3 width='100%' border=0><tr><td>ѡ��1��</td><td><input type='text' name='Votes1' style='width:300'></td></tr><tr><td>ѡ��2��</td><td><input type='text' name='Votes2' style='width:300'><INPUT TYPE='hidden' name='AutoValue' value='2'></td></tr></table></span></td></tr>"
	YxBBs.Stats="������ͶƱ"
	Caption="������ͶƱ"
	bbsCaption="<input name='Caption' type=text style='width:80%' maxlength=200>"
	Submiturl="Save.Asp?BoardID="&YxBBs.BoardID
End Sub

Sub Reply()
	Dim Rs,BbsID,TopicCaption
	if Id=0 Then YxBBs.Error("����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������")
	If Cint(YxBBs.ClassSetting(3))=0 Then YxBBs.Error(12)
	YxBBs.Stats="�ظ�����"
	IsTitle=False
	Set Rs=YxBBs.Execute("Select Caption,SqlTableID,IsLock From [YX_Topic] where TopicID="&ID)
	If Rs.Eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	ElseIf Rs(2) Then
		YxBBs.Error("�������Ѿ���������")
	Else
		TopicCaption=Rs(0)
		YxBBs.TB=Rs(1)
	End If
	Set Rs=Nothing
	bbsCaption=" <input type=hidden name='Caption' value='Re:"&TopicCaption&"'>"
	Caption="�ظ����⣺"&TopicCaption
	Submiturl="Save.Asp?Action=Reply&BoardID="&YxBBs.BoardID&"&TB="&YxBBs.TB&"&ID="&ID&"&page="&page&""
	BbsID=YxBBs.CheckNum(Request.QueryString("BbsID"))
	If BbsID<>"" Then
		Set Rs=YxBBs.Execute("select top 1 B.ReplyTopicID,B.TopicID,B.Name,B.AddTime,B.Content,B.BoardID,U.IsShow from [YX_Bbs"&YxBBs.TB&"] As B inner join [YX_User] As U on B.Name=U.Name where B.BbsID="&BbsID&" And B.IsDel=False")
			If Not Rs.Eof Then
				If Rs(1)<>ID And Rs(0)<> Id Then YxBBs.Error("����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������")
				If Rs(6) Then
					BbsCaption=BbsCaption &"<Font Color=Red>�������ݲ�������</Font>"
				Else
					BbsCaption=BbsCaption &" <Font Color=Red>���� "&Rs(2)&" �ķ���</Font>"
					Quote="<tr><td>&nbsp;<b>���õ����ݣ�<input name='QuoteName' type='hidden' value='[quote]����������  <font color=0000FF>"&Rs(2)&"</font>  ��( "&Rs(3)&" )�ķ��ԣ�"&chr(10)&chr(10)&"'></b><br>&nbsp;���Ա༭Ҫ���õ�����</td><td width='76%'>&nbsp;<TEXTAREA name='Quote'  cols=90 rows='5' style='WIDTH: 95%;'>"&YxBBs.Fun.StrLeft(QuoteCode(Rs(4)),500)&"</textarea></td>"
				End If
			End if
		Set Rs=Nothing
	End If
End Sub

Sub Edit()
	Dim Rs,BbsID,TopicIsLock,TopicRs,IsTop,Content
	BbsID=YxBBs.CheckNum(Request.QueryString("BbsID"))
	IF BbsID=0 Or ID=0 Then YxBBs.Error("����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������")
	Set Rs=YxBBs.Execute("Select BoardID,TopType,SqlTableID,IsLock,TitleType,face From [YX_Topic] where IsDel=False And TopicID="&ID)
	If Rs.Eof Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		TopicRs=Rs.GetRows(-1)
	End If
	Rs.Close
	'������ܶ���
	Set Rs=YxBBs.Execute("select BoardID,Name,AddTime,TopicID,Caption,Content,IsDel From [YX_Bbs"&TopicRs(2,0)&"] where IsDel=False And BbsID="&BbsID&"")
	Set Rs1=YxBBs.Execute("select face From [YX_Topic] where TopicID="&ID)
	If Rs.eof  Then
		YxBBs.Error("�����Ӳ����ڻ��Ѿ�ɾ����")
	Else
		If Not(TopicRs(1,0)=5 or TopicRs(1,0)=4) and TopicRs(0,0)<>YxBBs.BoardID Then YxBBs.Error("����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������")
		If lcase(YxBBs.MyName)<>lcase(rs("name")) Then
			If Cint(YxBBs.ClassSetting(18))=0 Then YxBBs.Error("��û�б༭�������ӵ�Ȩ�ޣ�")
		Else
			'�༭�Լ����ӵ�Ȩ��
			If Cint(YxBBs.ClassSetting(9))=0 Then YxBBs.Error("��û�б༭�Լ����ӵ�Ȩ�ޣ�")
		End If
		If YxBBs.ClassID>2 And Not YxBBs.IsBoardAdmin Then
			If YxBBs.BBSSetting(31)<>0 Then If lcase(YxBBs.MyName)=lcase(rs("name")) And DateDiff("s",Rs("AddTime")+YxBBs.BBSSetting(31)/1440,YxBBs.NowBbsTime)>0 Then YxBBs.Error("�㳬����������ͨ�û��༭�Լ����ӵ�ʱ�� (������<Font color=red>"&YxBBs.BBSSetting(31)&"</Font>������)")
			If lcase(YxBBs.MyName)=lcase(rs("name")) And TopicRs(3,0) Then YxBBs.Error("�������Ѿ���������")
		End If
		BbsCaption=rs("Caption")
				IF Rs("TopicID")=0 Then
			IsTitle=False
			bbsCaption="<input type=hidden name='Caption' value='"&BbsCaption&"'>�༭�ظ���"
		Else 
			bbsCaption="<input name='Caption' type=text value='"&bbsCaption&"' size=70 maxlength=200><input type=hidden name=OldTitleType value='"&TopicRs(4,0)&"'>"
		End If
		ThisFace=Rs1("Face")
		BbsContent=replaceUBB(rs("Content"))
	End if
	Rs.Close
	YxBBs.Stats="�༭����"
	Caption=YxBBs.Stats
	Submiturl="Save.Asp?Action=Edit&ID="&ID&"&BbsID="&BbsID&"&BoardID="&YxBBs.BoardID&"&TB="&YxBBs.TB&"&page="&page&""
End Sub

Function QuoteCode(str)
	Dim re,restr
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	restr="[COLOR=RED]���� �������ݲ������� ����[/COLOR]"
	re.Pattern="(\[DATE=(.[^\[]*)\])(.[^\[]*)(\[\/DATE\])"
	str=re.Replace(str,restr)
	re.Pattern="(\[SEX=*([0-1]*)\])(.[^\[]*)(\[\/SEX\])"
	str=re.Replace(str,restr)
	re.Pattern="(\[COIN=*([0-9]*)\])(.[^\[]*)(\[\/COIN\])"
	str=re.Replace(str,restr)		
	re.Pattern="(\[USERNAME=(.[^\[]*)\])(.[^\[]*)(\[\/USERNAME\])"
	str=re.Replace(str,restr)	
	re.Pattern="(\[GRADE=*([0-9]*)\])(.[^\[]*)(\[\/GRADE\])"
	str=re.Replace(str,restr)	
	re.Pattern="(\[MARK=*([0-9]*)\])(.[^\[]*)(\[\/MARK\])"		
	str=re.Replace(str,restr)
	re.Pattern="(\[BUYPOST=*([0-9]*)\])(.[^\[]*)(\[\/BUYPOST\])"
	str=re.Replace(str,restr)
	re.Pattern=vbcrlf&vbcrlf&vbcrlf&"(\[RIGHT\])(\[COLOR=(.[^\[]*)\])(.[^\[]*)(\[\/COLOR\])(\[\/RIGHT\])"
	str=re.Replace(str,"")
	re.Pattern="(<p align=right><font color=#0080FF>�������ӱ�(.*)�༭����</font></p>)" 
	str=re.Replace(str," ")
	Str = Replace(Str, chr(10), "")
	re.Pattern="(\[reply\])(.+?)(\[\/reply\])"
	str=re.Replace(str,restr)	
	set re=Nothing
	QuoteCode=str
End Function

Function replaceUBB(str)
	dim re
	If Str="" Then Exit Function
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	re.Pattern="(>)("&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern=vbNewLine
	Str=re.Replace(Str,"<br>")	
        re.Pattern="(<p align=right><font color=#0080FF>(.[^\[]*)</font></p>)"
        str=re.Replace(str," ")
	str=Replace(Str,"  ","&nbsp;&nbsp;")
	Set re=Nothing
	replaceUBB=str
End function

Function ShowMain()
	Dim Temp,Face,I,J,Rs_1
        
	Temp=Temp&"<form style='margin:0;' method=POST name='topicform' action='"&Submiturl&"' >"
	If IsTitle Then
		Temp=Temp&"<tr><td width='24%' height='30'><b>&nbsp;�������⣺</b><select name=title onChange=DoTitle(this.options[this.selectedIndex].value) style='font-size: 9pt'><option selected value="""">����</option><option value=[ԭ��]>[ԭ��]</option><option value=[ת��]>[ת��]</option><option value=[����]>[����]</option><option value=[����]>[����]</option><option value=[�Ƽ�]>[�Ƽ�]</option><option value=[����]>[����]</option><option value=[ע��]>[ע��]</option><option value=[��ͼ]>[��ͼ]</option><option value=[����]>[����]</option><option value=[����]>[����]</option></select></td><td>"
	Else
		Temp=Temp&Quote&"<tr><td colspan=2 height='30'>"
	End If
	
	If Action="reply" Then
		Temp=Temp&"&nbsp;"&bbsCaption&"</td></tr>"
	Else
		Temp=Temp&"&nbsp;"&bbsCaption&"&nbsp;&nbsp;<select name=""topicximoo""><option value=0>������Ŀ</option><option value=1>��ɫ��Ŀ</option><option value=2>��ɫ��Ŀ</option><option value=3>��ɫ��Ŀ</option></select></td></tr>"
	Temp=Temp&"<tr><td width='24%'>&nbsp;<b>ͼ�꣺</b></td><td>"


Face="&nbsp;<input name=face type=radio value=99 checked>��&nbsp;<script>for(i=1;i<=9;i=i+1) {document.write(""<input type=radio value=""+i+"" name=face><img border=0 src=images/face/""+i+"".gif>"")}</script>"
        End If
	Temp=Temp&Face&"</td></tr>"&bbsVote&"<tr><td>&nbsp;<b>�����ϴ����ͣ�</b><br>&nbsp;"&Join(Split(YxBBs.BoardSetting(1),"|"),",")&"</td><td>"
	
        If Int(Rs("GradeNum"))>Int(YxBBs.BBSSetting(36)) then
		Temp=Temp&"<table border=0 cellpadding=0 cellspacing=0><tr><td>&nbsp;<iframe Src='ViewFile.Asp?Action=Upload&BoardID="&YxBBs.BoardID&"' scrolling='no' frameborder='0' height='20' width='260'></iframe></td><td>[ÿ���������ϴ� <font color=blue>"&YxBBs.ClassSetting(11)&"</font> �� С�� <font color=blue>"&YxBBs.ClassSetting(12)&"</font> KB / ÿ��]</td></tr></table>"
        Else

               Temp=Temp&"���ĵȼ�δ�ﵽ <B><font Color=red>"&YxBBs.BBSSetting(36)&"��</font></B> �����ϴ��ļ���<br>"
        End if
	
	Temp=Temp&"</td></tr><tr><td valign='top'><table width='100%'  border='0' cellpadding='0' cellspacing='4'><tr><td colspan='2'><b>����������</b></td></tr>"
	Temp=Temp&"<tr><td rowspan='4' width=10 ></td></tr><tr><td><br><li><a href=javascript:Especial('[Reply]','����[\/Reply]')>�ظ��ɼ�</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:Especial('[Coin=1000]','����[\/Coin]')>��Ǯ�ɼ�</a></li><br><li><a href=javascript:Especial('[Grade=1]','����[\/Grade]')>�ȼ��ɼ�</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:Especial('[Username=�ǳ�]','����[\/Username]')>ָ������</a></li><br><li><a href=javascript:Especial('[Buypost=100]','����[\/Buypost]')>���ѿɼ�</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:Especial('[Date=2006-10-1]','����[\/Date]')>���ڿɼ�</a></li><br><li><a href=javascript:Especial('[Sex=1]','����[\/Sex]')>�Ա�ɼ�</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:Especial('[Login]','����[\/Login]')>��½�ɼ�</a></li></td></tr><tr><td><br><li>�����������������<font color=red>"&YxBBs.BoardSetting(3)/1024&"KB</font></li><br><br><li>�緢��ʧ��,��<font color=bule>Ctrl+v</font>���һ�����!</li></td></tr></table></td>"
	Temp=Temp&"<td width='76%' valign='top'><table border='0' cellpadding='0' cellspacing='3'  width='100%'><tr><td colspan='2'>"
        Temp=Temp&"<textarea name='content' cols='40' rows='2' style='display:none'>"&BbsContent&"</textarea><script type=""text/javascript"">editor.show()</script></tr></table></td></tr><tr><td height='40' colspan='2' align='center'>"
        If Action="edit" And (YxBBs.ClassID<=3) Then Temp= Temp & "<input name='EditChalk' type='checkbox' value='No'>�����±༭��� "
	If Cint(YxBBs.BoardSetting(0))=1 Then Temp=Temp&""&YxBBs.GCode&""
	Temp=Temp&"&nbsp;<input class=submit type=button value=OK_������ name='B1' onclick='submitform()'>&nbsp;&nbsp;&nbsp; <input class=submit type=button value='Ԥ  ����' onclick='preview()' name='B3'>&nbsp;&nbsp;&nbsp; <input class=submit type=reset value=NO_����д onclick='Goreset()' name=B2>&nbsp; [ �� Ctrl+Enter ֱ�ӷ��� ]" 
	Temp=Temp&"<font color=#D8D8D8></font></td></tr></form>"
	Call YxBBs.ShowTable(Caption,Temp)
	Response.Write"<form style='margin:0;' name='form1' action='See.asp?Action=preview' method='post' target='showgg'><input type='hidden' name='Caption' /><input type='hidden' name='content' /></form>"
	End Function
%>