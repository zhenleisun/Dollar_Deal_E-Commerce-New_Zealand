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
set rs=server.createobject("adodb.recordset")
rs.open "select * from vote order by id desc",conn,1,1
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/css.css" rel=stylesheet>
<head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form method="POST" action="setvote.asp">
<tr> 
<td colspan="5" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�̳�ͶƱ���� </font></b></td>
</tr>
<tr > 
<td width="10%" align="center" bgcolor="fbc2c2">ѡ��</td>
<td width="10%" align="center" bgcolor="fbc2c2">ID</td>
<td width="60%" align="center" bgcolor="fbc2c2">����</td>
<td width="10%" align="center" bgcolor="fbc2c2">�޸�</td>
<td width="10%" align="center" bgcolor="fbc2c2">ɾ��</td>
</tr>
<%
do while not rs.eof
%>
<tr > 
<td width="10%" align="center"> 
<input type="radio" value=<%=rs("ID")%><%if rs("IsChecked")=1 then%> checked<%end if%> name="Checked"></td>
<td width="10%" align="center"><%=rs("ID")%>��</td>
<td width="60%"><%=rs("Title")%>��</td>
<td width="10%" align="center"> 
<input onClick="javascript:window.open('modifyvote.asp?id=<%=rs("ID")%>','_self','')" type="button" value="�޸�" name="button1" style="font-family: ����; font-size: 9pt"></td>
<td width="10%" align="center"> 
<input onClick="javascript:window.open('delvote.asp?id=<%=rs("ID")%>','_self','')" type="button" value="ɾ��" name="button2" style="font-family: ����; font-size: 9pt"></td>
</tr>
<%
rs.movenext
loop
rs.close
%>
<tr> 
<td colspan=5 align=center bgcolor=#fbf4f4> 
<input type="submit" value="ѡ��ͶƱ��" name="submit" style="font-family: ����; font-size: 9pt"> <input onClick="javascript:window.open('addvote.asp','_self','')" type="button" value="�����ͶƱ" name="button" style="font-family: ����; font-size: 9pt">
</td>
</tr>
</form>
</table>
<%
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>
