<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%dim admin,password,verifycode
admin=replace(trim(request("admin")),"'","")
password=md5(replace(trim(request("password")),"'",""))
verifycode=replace(trim(request("verifycode")),"'","")
adminIP	   = Request.ServerVariables("Remote_Addr")
if admin="" or password="" then
dim ConGratUlatIons
set ConGratUlatIons=server.CreateObject("adodb.recordset")
ConGratUlatIons.Open "select * from cnhwwlog " ,conn,1,3
ConGratUlatIons.AddNew
ConGratUlatIons("adminuser")="��"
ConGratUlatIons("adminip")=adminIP
ConGratUlatIons("shijian")="��½ʧ�ܣ��������ĵ�½�������룡"
ConGratUlatIons.update
ConGratUlatIons.Close
set wangqu=nothing
response.Write "<script LANGUAGE='javascript'>alert('���Ĺ���ID����������');history.go(-1);</script>"
response.end
end if
if cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
set ConGratUlatIons=server.CreateObject("adodb.recordset")
ConGratUlatIons.Open "select * from cnhwwlog " ,conn,1,3
ConGratUlatIons.AddNew
ConGratUlatIons("adminuser")=admin
ConGratUlatIons("adminip")=adminIP
ConGratUlatIons("shijian")="��֤�벻��ȷ����¼ʧ�ܣ�"
ConGratUlatIons.update
ConGratUlatIons.Close
set ConGratUlatIons=nothing
response.Write "<script LANGUAGE='javascript'>alert('��������ȷ����֤�룡');history.go(-1);</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from cnhww where admin='"&admin&"' and password='"&password&"' " ,conn,1,1
if not(rs.bof and rs.eof) then
if password=rs("password") then
session("admin")=trim(rs("admin"))
session("flag")=int(rs("flag"))
session.Timeout=20
rs.Close
set rs=nothing
set ConGratUlatIons=server.CreateObject("adodb.recordset")
ConGratUlatIons.Open "select * from cnhwwlog " ,conn,1,3
ConGratUlatIons.AddNew
ConGratUlatIons("adminuser")=admin
ConGratUlatIons("adminip")=adminIP
ConGratUlatIons("shijian")="��¼�ɹ���"
ConGratUlatIons.update
ConGratUlatIons.Close
set wangqu=nothing
response.Redirect "index.asp"
else
set ConGratUlatIons=server.CreateObject("adodb.recordset")
ConGratUlatIons.Open "select * from cnhwwlog " ,conn,1,3
ConGratUlatIons.AddNew
ConGratUlatIons("adminuser")=admin
ConGratUlatIons("adminip")=adminIP
ConGratUlatIons("shijian")="������󣬵�¼ʧ�ܣ�"
ConGratUlatIons.update
ConGratUlatIons.Close
set ConGratUlatIons=nothing
response.write "<script LANGUAGE='javascript'>alert('�Բ��𣬵�½ʧ�ܣ�');history.go(-1);</script>"
end if
else
set ConGratUlatIons=server.CreateObject("adodb.recordset")
ConGratUlatIons.Open "select * from cnhwwlog " ,conn,1,3
ConGratUlatIons.AddNew
ConGratUlatIons("adminuser")=admin
ConGratUlatIons("adminip")=adminIP
ConGratUlatIons("shijian")="û�и��û�������¼ʧ�ܣ�"
ConGratUlatIons.update
ConGratUlatIons.Close
set ConGratUlatIons=nothing
response.write "<script LANGUAGE='javascript'>alert('�Բ��𣬵�½ʧ�ܣ�');history.go(-1);</script>"
end if
%>
