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
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">后台管理员设置</font></b></td>
</tr>
<tr align="center" > 
<td width="25%" bgcolor="fbc2c2">管理员</td>
<td width="25%" bgcolor="fbc2c2">密 码</td>
<td width="25%" bgcolor="fbc2c2">权 限</td>
<td width="25%" bgcolor="fbc2c2">操 作</td>
</tr>
		<%set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from cnhww order by flag",conn,1,1
		do while not rs.EOF%>
<form name="form1" method="post" action="saveadmin.asp?action=edit&id=<%=int(rs("adminid"))%>">
<tr align="center" > 
<td><input name="admin" type="text" size="16" value="<%=trim(rs("admin"))%>"></td>
<td><input name="password" type="text" size="16"></td>
<td><%select case rs("flag")
                case "1"
                response.Write "<input type=radio name=flag value=1 checked>管理&nbsp;<input name=flag type=radio value=2 >添加&nbsp;<input type=radio name=flag value=3>查看"
                case "2"
                response.Write "<input type=radio name=flag value=1>管理&nbsp;<input name=flag type=radio value=2 checked>添加&nbsp;<input type=radio name=flag value=3>查看"
				case "3"
				response.Write "<input type=radio name=flag value=1>管理&nbsp;<input name=flag type=radio value=2>添加&nbsp;<input type=radio name=flag value=3  checked>查看"
                end select%>
</td>
<td><input type="submit" name="Submit" value="修 改">
&nbsp;<a href="saveadmin.asp?id=<%=int(rs("adminid"))%>&action=del" onClick="return confirm('您确定要删除此用户吗？')"><font color=red>删除</font></a>
</td>
</tr>
</form>
		<%rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
</table><br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加管理员</font></b></td>
</tr>
<tr align="center" > 
<td width="25%" bgcolor="fbc2c2">管理员</td>
<td width="25%" bgcolor="fbc2c2">密 码</td>
<td width="25%" bgcolor="fbc2c2">权 限</td>
<td width="25%" bgcolor="fbc2c2">操 作</td>
</tr>
<form name="form1" method="post" action="saveadmin.asp?action=add">
<tr align="center" > 
<td><input name="admin2" type="text" size="16"></td>
<td><input name="password2" type="text" size="16"> </td>
<td><input type="radio" name="flag2" value="1">
管理 
<input name="flag2" type="radio" value="2" checked>
添加 
<input type="radio" name="flag2" value="3">
查看
</td>
<td>
<input type="submit" name="Submit2" value="添加管理员">
</td>
</tr>
</form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
