<!--#include file="conn.asp"-->
<%
if request.cookies("Cnhww")("username")="" then
response.write "<script language=javascript>alert('�Բ�������û�е�½��');window.close();</script>"
response.End
end if
username=request.cookies("Cnhww")("username")
bookid=request("bookid")
if bookid="" then
response.write "<script language=javascript>alert('�Բ�����û��ѡ����Ʒ��');window.close();</script>"
response.End
end if
Set rs_s=Server.CreateObject("Adodb.RecordSet")
rs_s.Open "Select * from products where bookid in ("&bookid&")",Conn,3,3
while not rs_s.eof 
if request.cookies("Cnhww")("reglx")=2 then 
	danjia=rs_s("vipjia")
else
	danjia=rs_s("huiyuanjia")
end if
kucun=rs_s("kucun")
bookname=rs_s("bookname")
if kucun<=0 then
response.write "<script language=javascript>alert('��ѡ������Ʒ��"&bookname&"����ʱȱ�����ܷŵ����ﳵ���ѡ��������Ʒ��');window.close();</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from orders where username='"&username&"' and bookid="&trim(rs_s("bookid"))&" and zhuangtai=7",conn,1,3
if rs.recordcount=1 then
if kucun<(rs("bookcount")+1) then
response.write "<script language=javascript>alert('��ѡ������Ʒ��"&bookname&"����ʱȱ�����ܷŵ����ﳵ���ѡ��������Ʒ��');window.close();</script>"
response.end
end if
rs("zonger")=(rs("bookcount")+1)*danjia
rs("bookcount")=rs("bookcount")+1
rs.update
rs.close
set rs=nothing
else
rs.close
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from orders",conn,1,3
rs.addnew
rs("bookid")=trim(rs_s("bookid"))
rs("username")=username
rs("zhuangtai")=7
rs("bookcount")=1
rs("zonger")=danjia
rs.update
rs.close
set rs=nothing
end if
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
response.Redirect "buy.asp?action=show"
%>
