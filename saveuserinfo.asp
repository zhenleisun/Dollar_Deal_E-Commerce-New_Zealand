<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%dim action,username
action=request.QueryString("action")
username=request.cookies("Cnhww")("username")
select case action
'//收货人信息
case "shouhuoxx"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [user] where username='"&username&"' ",conn,1,3
rs("shouname")=trim(request("shouname"))
rs("shengshi")=trim(request("shengshi"))
rs("reglx")=2	
rs("szSheng")=trim(request("szSheng"))
rs("szShi")=trim(request("szShi"))
rs("shouhuodizhi")=trim(request("shouhuodizhi"))
rs("youbian")=trim(request("youbian"))
rs("usertel")=trim(request("usertel"))
rs("songhuofangshi")=trim(request("songhuofangshi"))
rs("zhifufangshi")=trim(request("zhifufangshi"))
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('您的详细资料信息保存成功！');</script>"
response.redirect "myuser.asp?action=shouhuoxx"
response.End
case "userziliao"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&username&"'",conn,1,3
rs("useremail")=trim(request("useremail"))
if request("ifgongkai")="" then 
rs("ifgongkai")=0
else
rs("ifgongkai")=trim(request("ifgongkai"))
end if
rs("userzhenshiname")=trim(request("userzhenshiname"))
rs("sfz")=trim(request("sfz"))
rs("sex")=trim(request("shousex"))
rs("nianling")=trim(request("nianling"))
rs("szsheng")=trim(request("szsheng"))
rs("szshi")=trim(request("szshi"))
rs("shouhuodizhi")=trim(request("shouhuodizhi"))
rs("usertel")=trim(request("usertel"))
rs("youbian")=trim(request("youbian"))
rs("oicq")=trim(request("qq"))
rs("homepage")=trim(request("homepage"))
rs("content")=trim(request("content"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的个人资料修改成功！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
response.end
case "savepass"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [user] where username='"&username&"'",conn,1,3
if trim(request("userpassword"))<>"" then
rs("userpassword")=md5(trim(request("userpassword")))
end if
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('密码更改成功！');window.location.href='"&request.servervariables("http_referer")&"';</script>"
response.End
case "repass"
set rs=server.CreateObject("adodb.recordset")
rs.open "select userpassword from [user] where username='"&trim(request("username2"))&"'",conn,1,3
rs("userpassword")=md5(trim(request("userpassword2")))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的密码取回成功，请登陆！');history.go(-1);</script>"
end select
%>
