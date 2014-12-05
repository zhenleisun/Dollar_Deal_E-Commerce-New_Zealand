<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
dim action,dingdan,username
action=request.QueryString("action")
dingdan=request.QueryString("dan")
username=request.QueryString("username")
select case action
case "save"
if request("zhuangtai")<>"" then
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select zhuangtai from orders where dingdan='"&dingdan&"'",conn,1,3
	do while not rs.EOF
	old_zhuangtai=rs("zhuangtai")
		rs("zhuangtai")=request("zhuangtai")
		rs.Update
		rs.MoveNext
	loop
	rs.Close
	set rs=nothing
'conn.execute "update orders set zhuangtai="&request("zhuangtai")&" where dingdan='"&dingdan&"' "
end if
if request("zhuangtai")=4 then
fhsj=now()
conn.execute "update orders set fhsj=date() where dingdan='"&dingdan&"' "
end if
if cint(request("zhuangtai"))=5 and old_zhuangtai<>5 then
	jifen=0
	ifhuyuanka=0
		set rs2=server.CreateObject("adodb.recordset")
		rs2.Open "select vipid from charging",conn,1,1
		vipid=rs2("vipid")
		rs2.close
		set rs2=nothing
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select bookcount,bookid from orders where dingdan='"&dingdan&"'",conn,1,1
	while not rs.eof
		set rs2=server.CreateObject("adodb.recordset")
		rs2.Open "select bookid,yeshu from products where bookid="&rs("bookid"),conn,1,1
		jifen=jifen+rs("bookcount")*rs2("yeshu")
		rs2.close
		set rs2=nothing
		if rs("bookid")=cint(vipid) then 
			ifhuyuanka=1
		end if
		rs.MoveNext
	wend
	rs.Close
	'response.write ifhuyuanka&"'"&vipid
	'response.end
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select jifen,reglx,vipdate from [YX_User] where username='"&username&"'",conn,1,3
	rs("jifen")=rs("jifen")+jifen
	if ifhuyuanka=1 then 
		rs("reglx")=2
		if rs("vipdate")<>"" then 
		if rs("vipdate")<date then
		rs("vipdate")=date+365
		else
		rs("vipdate")=rs("vipdate")+365
		end if
		else
		rs("vipdate")=date+365
		end if
	end if
	rs.Update
	rs.Close
	set rs=nothing
	if ifhuyuanka=1 then 
		response.Write "<script language=javascript>alert('订单状态修改成功！客户本次购物获得积分:"&jifen&"，你本次购买了会员卡，恭喜你现在已经成为本站的VIP用户！！');history.go(-1);</script>"
	else
		response.Write "<script language=javascript>alert('订单状态修改成功！客户本次购物获得积分:"&jifen&"');history.go(-1);</script>"
	end if
else
	response.Write "<script language=javascript>alert('订单状态修改成功！');history.go(-1);</script>"
end if
case "del"
'删除时要判断状态，未收到货时要返还积分和库存的
set rs=server.createobject("adodb.recordset")
rs.open "select zhuangtai from orders where dingdan='"&dingdan&"' ",conn,1,1
if rs("zhuangtai")>7 then
rs.close
set rs=server.CreateObject("adodb.recordset")

rs.open "select * from orders where  dingdan='"&dingdan&"'",conn,1,1

while not rs.eof
	set rs_s=server.CreateObject("adodb.recordset")
	rs_s.open "select * from products where bookid="&rs("bookid"),conn,1,3
	rs_s("kucun")=rs_s("kucun")+rs("bookcount")
	rs_s("chengjiaocount")=rs_s("chengjiaocount")-rs("bookcount")
	rs_s.update
	rs_s.close
	set rs_s=nothing
rs.movenext
wend
rs.close
z_jifen=0
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ordersaward where  dingdan='"&dingdan&"'",conn,1,1
while not rs.eof
z_jifen=z_jifen+rs("jifen")
rs.movenext
wend
rs.close
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [YX_User] where username='"&request.Cookies("cnhww")("username")&"'",conn,1,3
rs("jifen")=rs("jifen")+z_jifen
rs.update
rs.close
set rs=nothing
else
rs.close
set rs=nothing
end if
conn.execute "delete from orders where dingdan='"&dingdan&"' "
response.Write "<script language=javascript>alert('订单删除成功！');window.close();window.opener.location.reload();</script>"
end select
%>
