<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
if request.cookies("Cnhww")("username")="" then
response.Redirect "user.asp"
response.End
end if
set cnhww=server.CreateObject("adodb.recordset")
cnhww.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"' ",conn,1,3
jf=cnhww("jifen")
yc=cnhww("yucun")
action=request("act")
if action="jifen" then
jifen=trim(request("jifen"))
if not isInteger(jifen) then
response.write"<script>alert(""非法换算!"");history.go(-1);</script>"
else
if jf-jifen<0 then 
response.write"<script>alert(""你们没有这么多积分可以换算!"");history.go(-1);</script>"
else
cnhww("jifen")=jf-jifen
cnhww("yucun")=yc+(jifen/2)
cnhww.update
response.redirect "user.asp"
end if
end if
end if
if action="cunkuan" then
cunkuan=trim(request("cunkuan"))
if not isInteger(cunkuan) then
response.write"<script language=javascript>alert(""非法换算!"");history.go(-1);</script>"
else
if yc-cunkuan<0 then 
response.write"<script language=javascript>alert(""你们没有这么多预存款可以换算!"");history.go(-1);</script>"
else
cnhww("jifen")=jf+(cunkuan*2)
cnhww("yucun")=yc-cunkuan
cnhww.update
response.redirect "user.asp"
end if
end if
end if
%>
