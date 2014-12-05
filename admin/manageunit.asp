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
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
    <form method="post" action="setunit.asp?action=update">
          <tr> 
            <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">单位设置</font></b></td>
          </tr>
      <tr align=center > 
        <td width="20%" bgcolor="fbc2c2">序 号</td>
        <td width="30%" bgcolor="fbc2c2">单位名称</td>
        <td width="30%" bgcolor="fbc2c2">排 序(必填项数字)</td>
        <td width="20%" bgcolor="fbc2c2">操 作</td>
      </tr>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "SELECT  * from unit  order by danweiorder",conn,1,1
if rs.recordcount=0 then 
%>
<tr> 
<td colspan="8" height=25 align="center" width="100%" > 还没有添加单位 </td>
</tr>
<%
else
    do while not rs.eof
%>
      <tr > 
        <td align=center><%=rs("ID")%></td>
        <td align=center> 
          <input type="hidden" name="danweiid" value="<%=rs("id")%>">
          <input class=text type="text" name="danweiname" size="10" value="<%=trim(rs("danweiname"))%>">
        </td>
        <td align=center> 
          <input class=text type="text" name="danweiorder" size="10" value="<%=rs("danweiorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" >
        </td>
        <td align=center> 
          <a href="delunit.asp?danweiid=<%=rs("id")%>">删除</a> 
        </td>
      </tr>
      <%
    rs.MoveNext
    Loop
%>
<tr> 
<td colspan="4"  align="center" width="100%"> 
<input type="submit" name="Submit2" value="保存修改" style="font-family: 宋体; font-size: 9pt">&nbsp;
<input type="reset" value="重新设定" style="font-family: 宋体; font-size: 9pt" name="重置" >
</td>
</tr>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</form>
</table><br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form method="post" action="setunit.asp?action=add">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加单位</font></b></td>
</tr>
<tr height=40 > 
<td width=20% align=center>添加单位</td>
<td width=30% align=center> 
<input class=text type="text" name="danweiname" size="10">
</td>
<td width=30% align=center> 
<input class=text type="text" name="danweiorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td width=20% align=center> 
<input type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt">
</td>
</tr>
</form>
</table>
</body>
</html>
