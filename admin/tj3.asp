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

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="manage.css" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
.STYLE1 {color: #000000}
-->
</style></head>
<BODY>

<%
action=request("day")
if action="" then action=date()
%>

<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
<tr class=backs><td height=18 class=td><span class="STYLE1">������ϸͳ�ƣ����ڣ�</span><%=action%><span class="STYLE1">��</span></td>
</tr>
<tr><td>
<table border="1" width="100%" cellSpacing=0 cellPadding=3 bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF">
  <tr>
    <td>����ʱ��</td>
    <td>����IP</td>
    <td>�Ӻδ�����</td>
  </tr>

<%
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from count_shop where day=#"&action&"# order by times desc"
rs.cursorlocation = 3     '�α�
rs.open sql,conn,1,1 
if rs.eof and rs.bof then
%>
  <tr>
    <td colspan=3>��ʱû��ͳ������</td>
  </tr>
<%
else


pages = 30			'ÿҳ��¼��
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
%>



  <tr>
	<td><%=rs("times")%></td>
	<td><%=rs("ip")%></td>
	<td><%
if rs("where")<>"" then
'response.write "<span title="&rs("where")&">"&lleft(rs("where"),50)&"</span>"
response.write "<a href='../gotourl.asp?"&rs("where")&"' title="&rs("where")&" target='_blank'>"&lleft(rs("where"),50)&"</span>"
else
response.write "���ղؼд򿪻�ֱ��������ַ"
end if
%>
	</td>
  </tr>

<%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop

end if

rs.close
set rs=nothing
set rs1=nothing
conn.close
set conn=nothing
%>

</table>
<%
call listpages()
%>
</td></tr>
<tr><td>
<input type="button" onClick="javascript:location.href='javascript:history.go(-1)';" value="������һҳ">
</td></tr>
</table>
</body></html>


<%
'��ҳ
sub listPages() 
if allpages <= 1 then exit sub 
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?day="&action&"&page=1>��ҳ</a> <a href="&request.ServerVariables("script_name")&"?day="&action&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?day="&action&"&page="&page+1&">��ҳ</a> <a href="&request.ServerVariables("script_name")&"?day="&action&"&page="&allpages&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ"
end sub
%>
