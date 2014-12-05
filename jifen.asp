<!--#include file="conn.asp"-->
<%
if request.Cookies("Cnhww")("username")="" then
response.write "<script language=javascript>alert('对不起，您还没有登陆！');window.close();</script>"
response.End
end if
dim bookid,username,action
action=request.QueryString("action")
username=trim(request.Cookies("Cnhww")("username"))
bookid=request.QueryString("id")
if action="del" then
'//删除收藏
conn.execute "delete from ordersaward where actionid="&request.QueryString("actionid")
if request("ll")=22 then
response.redirect "buy.asp?action=show"
else
response.redirect "user.asp?action=jifen"
end if
response.End
end if

if action="add" then
'判断用户是否已换这个奖品
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select count(*) as rec_count from ordersaward where bookid="&bookid&" and zhuangtai=7 and username='"&request.Cookies("Cnhww")("username")&"'",conn,1,1

if rs_s("rec_count")>0 then
	rs_s.close
	set rs_s=nothing
	response.write "<script language=javascript>alert('这个奖品您已经选了，请再选其他奖品！');history.go(-1);</script>"
	response.end
end if
rs_s.close
set rs_s=nothing
'判断用户的现有积分
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select jifen from YX_User where name='"&request.Cookies("Cnhww")("username")&"'",conn,1,1
jifen_xy=rs_s("jifen")
rs_s.close
set rs_s=nothing
'判断用户的已准备换的积分
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select jifen from ordersaward where username='"&request.Cookies("Cnhww")("username")&"' and zhuangtai=7",conn,1,1
jifen_yy=0
while not rs_s.eof
jifen_yy=jifen_yy+rs_s("jifen")
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
'判断奖品需要的积分
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select point from jiangpin where bookid="&bookid,conn,1,1
jifen_y=rs_s("point")
rs_s.close
set rs_s=nothing

if jifen_xy<(jifen_yy+jifen_y) then
	response.write "<script language=javascript>alert('您现有的积分不够兑换本奖品，请更换其他奖品！');history.go(-1);</script>"
	response.end
end if
'//商品，判断是否存在
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ordersaward",conn,1,3
rs.addnew
rs("bookid")=bookid
rs("username")=username
rs("zhuangtai")=7
rs("jifen")=jifen_y
rs("bookcount")=1
rs.update
rs.close
set rs=nothing
response.Redirect "user.asp?action=jifen"
end if
%>