<!--#include file="conn.asp"-->
<%dim dingdan,action
action=request.QueryString("action")
dingdan=request.QueryString("dan")
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
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select jifen,reglx,vipdate from [user] where username='"&request.cookies("Cnhww")("username")&"'",conn,1,3
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
		response.Write "<script language=javascript>alert('����״̬�޸ĳɹ��������ι����û���:"&jifen&"���㱾�ι����˻�Ա������ϲ�������Ѿ���Ϊ��վ��VIP�û�����');history.go(-1);</script>"
	else
		response.Write "<script language=javascript>alert('����״̬�޸ĳɹ��������ι����û���:"&jifen&"');history.go(-1);</script>"
	end if
else
	response.Write "<script language=javascript>alert('����״̬�޸ĳɹ���');history.go(-1);</script>"
end if
case "del"
set rs=server.CreateObject("adodb.recordset")
rs.open "select username,dingdan from orders where dingdan='"&dingdan&"' " ,conn,1,1
if request.cookies("Cnhww")("username")<>trim(rs("username")) then
response.Write "����Ȩɾ���˶���!"
response.End
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from orders where  dingdan='"&dingdan&"' and zhuangtai=1",conn,1,1
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
rs.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"'",conn,1,3
rs("jifen")=rs("jifen")+z_jifen
rs.update
rs.close
set rs=nothing
conn.execute "delete from orders where dingdan='"&dingdan&"' and zhuangtai=1"
conn.execute "delete from ordersaward where dingdan='"&dingdan&"'"
response.Write "<script language=javascript>alert('����ɾ���ɹ���');window.close();</script>"
end select
%>
