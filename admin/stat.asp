
<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if%>
<%if request.QueryString("action")="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dept",conn,1,3
rs("jsqtoday")=trim(request("jsqtoday"))
rs("jsq")=trim(request("jsq"))
rs.update
rs.close
set rs=nothing
response.write "<script language=javascript>alert('修改成功！');history.go(-1);</script>"
response.End
end if
%>
<html>
<head><title><%=webname%>访问统计</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dept",conn,1,1%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form name="form1" method="post" action="stat.asp?action=save">
<tr> 
<td colspan="2" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商城访问量设置</font></b></td>
</tr>
<tr >
<td align="center">今日商城访问量：</td>
<td style="PADDING-LEFT: 10px"> 
<input name="jsqtoday" type="text" id="458" size="28" value=<%=trim(rs("jsqtoday"))%>>
</td>
</tr>
<tr > 
<td align="center">商城的总访问量：</td>
<td style="PADDING-LEFT: 10px">
<input name="jsq" type="text" id="458url" size="28" value=<%=trim(rs("jsq"))%>>
</td>
</tr>
<tr > 
<td align="center"></td>
<td style="PADDING-LEFT: 10px"> 
<input type="submit" name="Submit" value="提交更改">
</td>
</tr>
</form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
