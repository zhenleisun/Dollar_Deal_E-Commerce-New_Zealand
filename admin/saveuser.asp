<!--#include file="conn.asp"-->
<!--#include file="../md5.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
dim userid,action
action=request.QueryString("action")
userid=request.QueryString("id")
if userid="" then userid=request("userid")
select case action
case "save"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from [user] where userid="&userid,conn,1,3
if trim(request("userpassword"))<>"" then rs("userpassword")=md5(trim(request("userpassword")))
rs("userzhenshiname")=trim(request("userzhenshiname"))
rs("useremail")=trim(request("useremail"))
rs("quesion")=trim(request("quesion"))
if trim(request("answer"))<>"" then rs("answer")=md5(trim(request("answer")))
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
if trim(request("vipdate"))<>"" then
    rs("vipdate")=trim(request("vipdate"))
end if
if trim(request("yucun"))<>"" then
rs("yucun")=trim(request("yucun"))
else
rs("yucun")=0
end if
rs("reglx")=trim(request("reglx"))
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('操作成功!');history.go(-1);</script>"
case "del"
conn.execute "delete from [user] where userid in ("&userid&") "
conn.execute "delete from orders where userid in ("&userid&")"
conn.execute "delete from ordersaward where userid in ("&userid&")"
conn.execute "delete from history where userid in ("&userid&")"
'response.Redirect "manageuser.asp"
response.Redirect request.servervariables("http_referer")
end select
%>
