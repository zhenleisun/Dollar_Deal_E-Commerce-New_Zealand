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

<script language=javascript src=../include/mouse_on_title.js></script>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="manage.css" type="text/css">
</head>
<BODY background="../images/admin/back.gif" bgcolor="#E7E7E7">

	<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">

	<tr class=backs><td class=td height=18>�ռ���</td></tr>
	
<%                                       
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from sms where (name='����Ա' or name='admin') and zuti='0' order by riqi desc"
rs.open sql,conn,1,3

if rs.eof and rs.bof then 
response.write "<tr><td>�ռ�����û����Ϣ��</td></tr>"
else

pages = 10			'ÿҳ��¼��
rs.pageSize = pages
allPages = rs.pageCount		'����һ���ֶܷ���ҳ
page = Request("page")
'if������ڻ������Ŵ���
if isEmpty(page) or clng(page) < 1 then
page = 1
elseif clng(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page

do while not rs.eof and pages>0
neirong=rs("neirong")
riqi=rs("riqi")
isnew=rs("isnew")
fname=rs("fname")
id=rs("id")
if pages<10 then response.write "<tr><td height=15></td></tr>"
%>
<tr><td onMouseOver="bgColor='#FFFFFF'" onMouseOut="bgColor='#e7e7e7'">
<span alt=�ظ�����Ϣ style="cursor:hand" onclick="{location.href='hand2.asp?back=hand1&id=<%=id%>';}"><img src=../images/small/m_replyp.gif border=0 align=center></span>&nbsp;&nbsp;
<span alt=ת����������Ա style="cursor:hand" onclick="{location.href='hand2.asp?back=hand1&fw=yes&id=<%=id%>';}"><img src=../images/small/m_fw.gif border=0 align=center></span>&nbsp;&nbsp;
<span alt=����ɾ������Ϣ style="cursor:hand" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��������Ϣ�� ')){location.href='hand4.asp?delid=<%=id%>';}}"><img src=../images/small/m_delete.gif border=0 align=center></span>&nbsp;&nbsp;
�����ˣ�<font color=red><%=fname%></font> &nbsp; &nbsp; &nbsp; ����ʱ�䣺<%=riqi%>

<hr color=#e7e7e7><BR>
<%=replace(neirong,vbCRLF,"<BR>")%><br>&nbsp;
</td></tr>
<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop
response.write "</table>"
call listpages()
end if
rs.close
set rs=nothing


'��ҳ
sub listPages() 
'if allpages <= 1 then exit sub 
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?page=1>��ҳ</a> <a href="&request.ServerVariables("script_name")&"?page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?page="&page+1&">��ҳ</a> <a href="&request.ServerVariables("script_name")&"?page="&allpages&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ"
end sub
%>


</td></tr></table>
</body>
</html>

