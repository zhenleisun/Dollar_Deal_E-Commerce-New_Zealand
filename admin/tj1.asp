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
body,td,th {
	color: #000000;
}
body {
	background-color: #FFFFFF;
}
.STYLE1 {color: #000000}
-->
</style></head>
<BODY>
<%
action = request.form("del")
if action="ok" then

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from count_total"
rs.open sql,conn,1,3
rs("total")=0		'�ܷ�������0
rs("yesterday")=0	'�����������0
rs("today")=0		'�����������0
rs.update
conn.execute "delete from count_shop"
rs.close
set rs=nothing

response.write "<script language='javascript'>"
response.write "alert('ͳ�������ѱ������ϵͳ�������ڿ�ʼ����ͳ�ơ�');"
response.write "</script>"
end if

conn.execute("delete from count_online where datediff('h',time,now())>1")

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from count_total"
rs.open sql,conn,1,3
total=rs("total")		'�ܷ�����
yesterday=rs("yesterday")	'���������
today=rs("today")		'���������

set rs = server.createobject("adodb.recordset")
sql = "select * from count_online"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
online=rs.RecordCount		'��������
else 
online=1
end if

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop order by day asc"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
total_ip=rs.RecordCount		'��IP������
firstday=rs("day")		'firstday����ʼ���������
else 
total_ip=0
firstday=date()
end if
per_ip=int(total_ip/(date()-firstday+1))	'ƽ��ÿ��IP������
if per_ip<1 then per_ip=0	
per=int(total/(date()-firstday+1))	'ƽ��ÿ�������
if per<1 then per=0

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop where day=#"&date()&"#"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
today_ip=rs.RecordCount		'������IP������
else 
today_ip=0
end if
set rs=nothing

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop where day=#"&date()-1&"#"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
yesterday_ip=rs.RecordCount		'������IP������
else 
yesterday_ip=0
end if
set rs=nothing

conn.close
set conn=nothing
%>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
<form action="tj1.asp" name="tongji" method=post>
<tr class=backs><td height=18 class=td STYLE1>��վ����ͳ�ƻ�������</td>
</tr>
<tr><td>
<table border="1" width="100%" cellSpacing=0 cellPadding=3 bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF">
  <tr>
    <td width="22%">ͳ����ʼ����</td>
    <td width="28%"><%=firstday%></td>
    <td width="22%">ͳ�ƽ�ֹ����</td>
    <td width="28%"><%=date()%>����<%=date()-firstday+1%>�죩</td>
  </tr>
  <tr>
    <td>�ܷ�����</td>
    <td><%=total%></td>
    <td>ƽ��ÿ�շ�����</td>
    <td><%=per%></td>
  </tr>
  <tr>
    <td>��IP������</td>
    <td><%=total_ip%></td>
    <td>ƽ��ÿ��IP������</td>
    <td><%=per_ip%></td>
  </tr>
  <tr>
    <td>���������</td>
    <td><%=yesterday%></td>
    <td>����IP������</td>
    <td><%=yesterday_ip%> &nbsp; <a href=tj3.asp?day=<%=date()-1%>>�鿴��ϸ</a></td>
  </tr>
  <tr>
    <td>���������</td>
    <td><%=today%></td>
    <td>����IP������</td>
    <td><%=today_ip%> &nbsp; <a href=tj3.asp?day=<%=date()%>>�鿴��ϸ</a></td>
  </tr>
  <tr>
    <td>��ǰ��������</td>
    <td><%=online%> ��</td>
    <td>ͳ����������</td>
    <td>
<input type=hidden name="del" value="ok">
<input type="submit" name="action" value="ִ��" onClick="{if(confirm('���ȫ��ͳ�����ݣ����¿�ʼͳ�ơ��˲����޷��ָ�����ȷ��Ҫ������')){this.document.tongji.submit();return true;}return false;}">
    </td></tr>
</table>
</td></tr>
</form>
</table>
</body></html>
