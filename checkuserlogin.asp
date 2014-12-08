<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%dim userpassword,comeurl,useremail
useremail=replace(trim(request("username")),"'","")
userpassword=md5(replace(trim(request("userpassword")),"'",""))
if useremail="" or userpassword="" then
response.write "<script LANGUAGE='javascript'>alert('Your email or password is empty.');history.go(-1);</script>"
response.end
end if
    %>
<%
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [user] where UserEmail='"&useremail&"' and userpassword='"&userpassword&"' " ,conn,1,3
if not(rs.bof and rs.eof) then
if userpassword=rs("userpassword") then
response.cookies("Cnhww")("useremail")=trim(request("username"))
response.cookies("Cnhww")("reglx")=rs("reglx")
response.cookies("Cnhww")("jifen")=rs("jifen")
response.cookies("Cnhww")("jiaoyijine")=rs("jiaoyijine")
rs("lastlogin")=now()
rs("logins")=rs("logins")+1
rs("userlastip")=Request.ServerVariables("REMOTE_ADDR")
rs.Update
rs.Close
set rs=nothing
useremail=trim(request("username"))
conn.execute("delete from orders where useremail='"&useremail&"' and zhuangtai=7")
conn.execute("delete from ordersaward where useremail='"&useremail&"' and zhuangtai=7")
if request("linkaddress")="" then
response.redirect request.servervariables("http_referer")
else
response.redirect request("linkaddress")
end if
else
response.write "<script LANGUAGE='javascript'>alert('Sorry, your email or password is incorrect¡ê?');history.go(-1);</script>"
end if
else
response.write "<script LANGUAGE='javascript'>alert('Sorry! Your email or password is incorrect¡ê?');history.go(-1);</script>"
end if
sub loginok()
response.Write "Welcome you <font color=red>"&request.cookies("Cnhww")("username")&"</font>"
response.redirect "index.asp"
end sub
%>
