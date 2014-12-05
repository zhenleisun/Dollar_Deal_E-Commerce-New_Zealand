
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
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">

</head>
<body>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">统计报表</font></b></td>
</tr>
<tr> 
<td bgcolor="fbf4f4" > 
<table width="400" border="0" align="center" cellpadding="2" cellspacing="1">
<form name="form1" method="post" action="reportform.asp">
<tr> 
<td>开始日期：</td>
<td>
<select name="year1">
<%for i=2006 to 2100%>
<option value="<%=i%>"  <%if year(now())=i then%>selected<%end if%>><%=i%></option>
<%next%>
</select>
年 
<select name="month1">
<%for i=1 to 12%>
<option value="<%=i%>"  <%if year(date())=i then %>selected<% end if %>><%=i%></option>
<%next%>
</select>
月 
<select name="day1">
<%for i=1 to 31%>
<option value="<%=i%>"  ><%=i%></option>
<%next%>
</select>
日 </td>
</tr>
<tr> 
<td width="70">结束日期：</td>
<td width="278">
<select name="year2">
<%for i=2006 to 2100%>
<option value="<%=i%>" ><%=i%> </option>
<%next%>
</select>
年 
<select name="month2">
<%for i=1 to 12%>
<option value=<%=i%>  ><%=i%> </option>
<%next%>
</select>
月 
<select name="day2">
<%for i=1 to 31%>
<option value=<%=i%>  <%if year(date())=i then%>selected<%end if%>><%=i%> </option>
<%next%>
</select>
日 </td>
</tr>
<tr> 
<td colspan="2" align="center"> 
<input type="submit" name="Submit" value="统 计">
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
