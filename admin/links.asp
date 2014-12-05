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

<%set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from links order by linkidorder",conn,1,1
		dim i
		i=rs.recordcount%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商品查看与修改</font></b></td>
</tr>
<tr bgcolor="#fbc2c2"> 
<td width="30%" align="center">网站名称</td>
<td width="30%" align="center">网站地址</td>
<td width="20%" align="center">排 序</td>
<td width="20%" align="center">操 作</td>
</tr>
			<%if rs.eof and rs.bof then
			response.write "还没有数据，请添加！"
			else
			do while not rs.eof%>
<tr bgcolor="#fbc2c2"> 
<form name="form1" method="post" action="savelinks.asp?action=edit&id=<%=rs("linkid")%>">
<td align="center" bgcolor="#fbf4f4"><input name="linkname" type="text" id="linkname" value="<%=trim(rs("linkname"))%>" size="16">
</td>
<td align="center" bgcolor="#fbf4f4">
<input name="linkurl" type="text" id="linkurl" value="<%=trim(rs("linkurl"))%>" size="26">
</td>
<td align="center" bgcolor="#fbf4f4">
<input name="linkidorder" type="text" id="linkidorder" value=<%=rs("linkidorder")%> size="3">
</td>
<td align="center" bgcolor="#fbf4f4">
<input type="submit" name="Submit" value="修 改">
&nbsp;<a href=savelinks.asp?action=del&id=<%=rs("linkid")%>><font color="#FF0000">删除</font></a> 
</td>
</form>
</tr>
<%rs.movenext
		  loop
		  end if
		  rs.close
		  set rs=nothing%>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form name="form2" method="post" action="savelinks.asp?action=add">
<tr>
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加合作伙伴</font></b></td>
</tr>
<tr bgcolor="#fbc2c2"> 
<td width="30%" align="center">网站名称 </td>
<td width="30%" align="center">网站地址 </td>
<td width="20%" align="center">排 序 </td>
<td width="20%" align="center">操 作 </td>
</tr>
<tr bgcolor="#fbc2c2"> 
<td align="center" bgcolor="#fbf4f4">
<input name="linkname1" type="text" id="linkname1" size="16">
</td>
<td align="center" bgcolor="#fbf4f4">
<input name="linkurl1" type="text" id="linkurl1" size="26">
</td>
<td align="center" bgcolor="#fbf4f4">
<input name="linkidorder1" type="text" id="linkidorder1" value=<%=i+1%> size="3">
</td>
<td align="center" bgcolor="#fbf4f4">
<input type="submit" name="Submit2" value="添加合作伙伴">
</td>
</tr>
</form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
