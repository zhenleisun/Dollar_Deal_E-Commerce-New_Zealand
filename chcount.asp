<!--#include file="conn.asp"-->
<%
actionid=request("actionid")
if actionid="" then
response.write "<script language=javascript>alert('对不起，您没有选择商品！');window.close();</script>"
response.End
end if
for i=1 to request.form("actionid").count
if request.form("bookcount")(i)<=0 then
bookcount=1
else
bookcount=request.form("bookcount")(i)
end if
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from products where bookid="&request.form("bookid")(i),conn,1,1
if request.cookies("Cnhww")("reglx")=2 then 
	danjia=rs_s("vipjia")
else
	danjia=rs_s("huiyuanjia")
end if
kucun=rs_s("kucun")
bookname=rs_s("bookname")
rs_s.close
set rs_s=nothing
if kucun<cint(bookcount) then
response.write "<script language=javascript>alert('你选购的商品“"&bookname&"”库存不足，不能修改数量，请选购更换其它商品！');window.location.href='buy.asp?action=show';</script>"
response.end
end if
conn.execute("update orders set bookcount="&bookcount&",zonger="&danjia*bookcount&" where actionid="&request.form("actionid")(i))
next
response.Redirect "buy.asp?action=show"
%>
