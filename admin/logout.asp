<%session("admin")=""
session("flag")=""
response.Write "<script language='javascript'>alert('您已成功注销登陆')</script>"
response.redirect "../index.asp"
%> 
