
<%if session("admin")="" then
response.Redirect "login.asp"
end if
%>
<html>
<head>
<title>购物系统--管理中心</title>
</head>
<frameset framespacing="0" border="false" cols="180,*" frameborder="0"> 
  <frame name="left"  scrolling="auto" marginwidth="0" marginheight="0" src="menu.asp">
  <frame name="right" scrolling="auto" src="admin.asp">
</frameset>
<noframes>
</noframes> 
</html>
