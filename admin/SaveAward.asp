<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
<%
if request("bookname")="" or request("shichangjia")="" or request("point")="" then
response.Write "<script language='javascript'>alert('������ȷ�ķ�ʽ��ӻ��޸Ľ�Ʒ��');history.go(-1);</script>"
response.End
end if
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
dim bookdate,dazhe
'dazhe=round(request("huiyuanjia")/request("shichangjia"),2)
'if request("bookdateyear")<>"" then
'bookdate=trim(request("bookdateyear"))&"��"&trim(request("bookdatemonth"))&"��"
'else
'bookdate=""
'end if
dim action,bookid
bookid=request.QueryString("id")
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from award",conn,1,3
rs.AddNew
rs("bookname")=trim(request("bookname")) '��Ʒ����
rs("guige")=trim(request("guige")) '��Ʒ���
rs("shichangjia")=trim(request("shichangjia"))  '�г���
rs("point")=trim(request("point"))  '����
rs("bookpic")=trim(request("bookpic"))  'СͼƬ��ַ
rs("bookpic2")=trim(request("bookpic2")) '��ͼƬ��ַ
rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  '���
rs("adddate")=now() '��������
if request("xianshi")=1 then  '�Ƽ�
rs("xianshi")=1
else
rs("xianshi")=0
end if
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('��ӳɹ���');window.location.href='AddAward.asp';</script>"
response.End
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from award where bookid="&bookid,conn,1,3
rs("bookname")=trim(request("bookname")) '��Ʒ����
rs("guige")=trim(request("guige")) '��Ʒ���
rs("shichangjia")=trim(request("shichangjia"))  '�г���
rs("point")=trim(request("point"))  '����
rs("bookpic")=trim(request("bookpic"))  'СͼƬ��ַ
rs("bookpic2")=trim(request("bookpic2")) '��ͼƬ��ַ
rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  '���
'rs("adddate")="bookdate"    '��������
if request("xianshi")=1 then  '�Ƽ�
rs("xianshi")=1
else
rs("xianshi")=0
end if
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('�޸ĳɹ���');window.location.href='ManageAward.asp';</script>"
response.End
end select
%>