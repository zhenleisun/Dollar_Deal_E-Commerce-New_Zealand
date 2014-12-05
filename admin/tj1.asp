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
rs("total")=0		'总访问量置0
rs("yesterday")=0	'昨天访问量置0
rs("today")=0		'今天访问量置0
rs.update
conn.execute "delete from count_shop"
rs.close
set rs=nothing

response.write "<script language='javascript'>"
response.write "alert('统计数据已被清除，系统将从现在开始重新统计。');"
response.write "</script>"
end if

conn.execute("delete from count_online where datediff('h',time,now())>1")

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from count_total"
rs.open sql,conn,1,3
total=rs("total")		'总访问量
yesterday=rs("yesterday")	'昨天访问量
today=rs("today")		'今天访问量

set rs = server.createobject("adodb.recordset")
sql = "select * from count_online"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
online=rs.RecordCount		'在线人数
else 
online=1
end if

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop order by day asc"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
total_ip=rs.RecordCount		'总IP访问量
firstday=rs("day")		'firstday：开始计算的日期
else 
total_ip=0
firstday=date()
end if
per_ip=int(total_ip/(date()-firstday+1))	'平均每天IP访问量
if per_ip<1 then per_ip=0	
per=int(total/(date()-firstday+1))	'平均每天访问量
if per<1 then per=0

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop where day=#"&date()&"#"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
today_ip=rs.RecordCount		'今天总IP访问量
else 
today_ip=0
end if
set rs=nothing

set rs = server.createobject("adodb.recordset")
sql = "select * from count_shop where day=#"&date()-1&"#"
rs.open sql,conn,1,3 
if not (rs.eof and rs.bof) then
yesterday_ip=rs.RecordCount		'昨天总IP访问量
else 
yesterday_ip=0
end if
set rs=nothing

conn.close
set conn=nothing
%>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#333333" cellspacing="0" cellpadding="2">
<form action="tj1.asp" name="tongji" method=post>
<tr class=backs><td height=18 class=td STYLE1>网站访问统计基本数据</td>
</tr>
<tr><td>
<table border="1" width="100%" cellSpacing=0 cellPadding=3 bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bordercolor="#FFFFFF">
  <tr>
    <td width="22%">统计起始日期</td>
    <td width="28%"><%=firstday%></td>
    <td width="22%">统计截止日期</td>
    <td width="28%"><%=date()%>（共<%=date()-firstday+1%>天）</td>
  </tr>
  <tr>
    <td>总访问量</td>
    <td><%=total%></td>
    <td>平均每日访问量</td>
    <td><%=per%></td>
  </tr>
  <tr>
    <td>总IP访问量</td>
    <td><%=total_ip%></td>
    <td>平均每日IP访问量</td>
    <td><%=per_ip%></td>
  </tr>
  <tr>
    <td>昨天访问量</td>
    <td><%=yesterday%></td>
    <td>昨天IP访问量</td>
    <td><%=yesterday_ip%> &nbsp; <a href=tj3.asp?day=<%=date()-1%>>查看明细</a></td>
  </tr>
  <tr>
    <td>今天访问量</td>
    <td><%=today%></td>
    <td>今天IP访问量</td>
    <td><%=today_ip%> &nbsp; <a href=tj3.asp?day=<%=date()%>>查看明细</a></td>
  </tr>
  <tr>
    <td>当前在线人数</td>
    <td><%=online%> 人</td>
    <td>统计数据置零</td>
    <td>
<input type=hidden name="del" value="ok">
<input type="submit" name="action" value="执行" onClick="{if(confirm('清除全部统计数据，重新开始统计。此操作无法恢复，您确定要继续吗？')){this.document.tongji.submit();return true;}return false;}">
    </td></tr>
</table>
</td></tr>
</form>
</table>
</body></html>
