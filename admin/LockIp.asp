<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<div align=center><font size=80 color=red><b>您没有此项目管理权限！</b></font></div>"
response.End
end if
end if
%>
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
  <tr> 
    <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商城锁定IP限制设置</font></b></td>
  </tr>
  <tr  style="PADDING-LEFT: 20px"> 
    <td >  
	
	
	
	
	
	
	
	<%
rows=request("rows")
if rows="" or rows=0 then rows=10
action=request("ok")
if action="" then 
Set rs = conn.Execute("select * from webinfo") 
%>
<table width="98%" border="1"  style="border-collapse: collapse; border-style: dotted; border-width: 0px"  bordercolor="#b1bfee" cellspacing="0" cellpadding="2"><form action=lockip.asp method=post name=setup>
  <tr> 
    <td width=20% align=right>是否开启访问限制&nbsp;</td>
    <td > <input type="radio" name="lockip" value="1" <%if rs("lockip")="1" then%>checked<%end if%>>
      是 
      <input type="radio" name="lockip" value="0" <%if rs("lockip")<>"1" then%>checked<%end if%>>
      否</td>
  </TR>
  <tr>
    <td width=20% align=right>锁定IP列表 &nbsp;</td>
    <td><TEXTAREA alt="请输入限制访问本站的IP地址<br>如：205.158.40.15<br>每个IP地址占一行" NAME="IP" ROWS="<%=rows%>" COLS="25" style="overflow:auto;">
<%
if rs("IP")<>"" then response.write Replace(rs("IP"),"@",vbCRLF)
%>
</TEXTAREA></td>
  </TR>
  <tr>
    <td colspan=2 width="20" ><INPUT name="ok" TYPE="hidden" value="ok">
      <INPUT name=action TYPE="submit" value="保存设置">
</table>
</form>
<font color=red>提示：</font>若输入IP地址的一部分，如：220.50，那么任何包含220.50的IP地址都将禁止访问本站。 
<DIV style="position:absolute; top:150; left:400"> 
<table border=0 width=100><tr><td width=50% align=center>
<form action=lockip.asp method=post name=setup2>
<INPUT name="rows" TYPE="hidden" value="<%=rows+10%>">
          <input type=image src=images/jia.gif alt="增加编辑区域" width="20" height="20" border=0>
</form>
</td>
<td width=100 align=center>
<form action=lockip.asp method=post name=setup3>
<INPUT name="rows" TYPE="hidden" value="<%=rows-10%>">
          <input type=image src=images/jian.gif alt="减少编辑区域" width="20" height="20" border=0>
</form>
</td></tr></table>
</td></tr>
</table>
</DIV> 
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
if action="ok" then
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from webinfo"
rs.open sql,conn,1,3
rs("lockip")=request.form("lockip")
ip=replace(request.form("ip")," ","")
ip=replace(ip,vbCRLF,"@")
for i=1 to 5
ip=replace(ip,"@@","@")
next
if right(ip,1)="@" then ip=left(ip,len(ip)-1)
if left(ip,1)="@" then ip=right(ip,len(ip)-1)
rs("IP")=ip
rs.update
url="lockip.asp"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script language='javascript'>"
response.write "alert('操作成功，您设置的信息已保存！');"
response.write "location.href='"&url&"';"
response.write "</script>"
end if
%>
	
	
	
	
	
	
	
	
	
	
	
	
	
</td>
  </tr>
  <td width="25%"> 
</table>
<!--#include file="foot.asp"-->
</body>
</html>

