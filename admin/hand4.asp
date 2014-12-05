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
<script language=javascript src=../include/mouse_on_title.js></script>
<%

	if request("delid")<>"" then	
	sql= "delete from sms where id="&request("delid")
	url="hand1.asp"
	end if

	if request("delid3")<>"" then	
	sql= "delete from sms where id="&request("delid3")
	url="hand3.asp"
	end if

	if request("delid5")<>"" then	
	sql= "delete from sms where id="&request("delid5")
	url="hand5.asp"
	end if

	if request("action")="all" then
	sql= "delete from sms"
	url="hand4.asp"
	end if

	if request("action")="w" then      
	sql="delete from sms where riqi<now()-7"
	url="hand4.asp"
	end if

	if request("action")="m" then
	sql="delete from sms where riqi<now()-30"
	url="hand4.asp"
	end if

if sql<>"" then
	conn.execute (sql)
	response.write "<script language='javascript'>"
	response.write "alert('所选站内消息删除成功！');"
	response.write "location.href='"&url&"';"
	response.write "</script>"
	response.end
else
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="manage.css" type="text/css">
</head>
<BODY background="../images/admin/back.gif" bgcolor="#C0C0C0">

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=hand4.asp method=post name="hand">
	<tr class=backs><td colspan=2 class=td height=18>删除站内消息</td></tr>

	<tr><td width=20% align=right height="18">删除&nbsp;</td><td>
	<input type=radio value="w" name="action" checked>一周前 
	<input type=radio value="m" name="action">一月前 
	<input type=radio value="all" name="action">所有消息 
	</td></tr>

	<tr><td colspan=2>
	<INPUT name=del TYPE="submit" value="立即删除" onclick="{if(confirm('此操作不可恢复，您确认要删除所选信息吗？\n\n单击确定继续，单击取消返回。')){this.document.hand.submit();return true;}return false;}"></td></tr>
	</form>
	</table>


</body>
</html>
<%
end if
%>