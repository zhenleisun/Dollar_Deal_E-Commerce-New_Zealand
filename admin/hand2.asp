<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if%>


<script language=javascript src=../include/mouse_on_title.js></script>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="manage.css" type="text/css">
</head>
<BODY background="../images/admin/back.gif" bgcolor="#C0C0C0">
<%
action=request("send")
if action<>"ok" then 
name=request("name")
id=request("id")
if id<>"" then
Set rs = conn.Execute("select * from sms where id="&id) 
name=rs("fname")
huifu=vbcrlf&vbcrlf&vbcrlf&vbcrlf&"======="&rs("fname")&"��"&rs("riqi")&"���͵���Ϣԭ��======="&vbcrlf&replace(rs("neirong"),"<font color=darkgray>","")
end if
%>

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
	<form action=hand2.asp method=post name="hand">
	<tr class=backs><td colspan=2 class=td height=18>����վ����Ϣ</td></tr>
	<tr><td width=20% align=right height="18">�� �� ˭&nbsp;</td><td>
	<input type=text 
	<%
	if request("fw")<>"yes" then response.write "value='"&name&"'"
	%>"
	name="name" size=18 maxlength=26> 
	<img src=../images/memo.gif border=0 alt="����д��վע���Ա��ID<br>" width="16" height="16"><font color=red>С�ؼ������͸������ˣ�����д��д��ALL</font></td></tr>
	<tr><td width=20% align=right height="18">��Ϣ����&nbsp;</td><td>
	<textarea rows="10" name="neirong" style="width=95%; overflow:auto;"><%=huifu%></textarea>
	<tr><td colspan=2>
	<INPUT name="send" TYPE="hidden" value="ok">
	<INPUT name="back" TYPE="hidden" value="<%=request("back")%>">
	<INPUT name=action TYPE="submit" value="������Ϣ"></td></tr>
	</form>
	</table>
<%

set rs=nothing
end if

if action="ok" then
	if trim(request.form("name"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û����д�ռ��ˣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if request.form("neirong")=""  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û����дҪ���͵����ݣ�');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if len(request.form("neirong"))>500  then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�����д������̫���ˣ�����������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if  request("name")<>"ALL" then
	name=request("name")
	 set rsv=server.createobject("adodb.recordset")
	  exec="SELECT * from [YX_User] where name='"&trim(request("name"))&"'"
	 rsv.open exec,conn,1,1
	if rsv.eof and rsv.bof then
	response.write "<script language='javascript'>"
	response.write "alert('�����ˣ�û���ҵ�"&request("name")&"��Ա������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if
	end if


	sql="select * from sms"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	  	rs("name")=trim(request.form("name"))
		neirong=request.form("neirong")
		neirong=replace(neirong,"=======","<font color=darkgray>======")
	  	rs("neirong")=neirong
	  	rs("riqi")=now()
	  	rs("fname")="admin"

	rs.update	
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

	where=request("where")
	if where="" then
	url="hand2.asp"
	else
	url=where
	end if

	response.write "<script language='javascript'>"
	response.write "alert('վ����Ϣ���ͳɹ���');"
	if request("back")="user" then
	response.write "location.href='user1.asp?action=useredit&id="&request("name")&"';"
	elseif request("back")="hand1" then
	response.write "location.href='hand1.asp';"
	elseif request("back")="hand3" then
	response.write "location.href='hand3.asp';"
	elseif request("back")="hand5" then
	response.write "location.href='hand5.asp';"
	else
	response.write "location.href='hand2.asp';"
	end if
	response.write "</script>"
	response.end
end if%>

</body>
</html>

