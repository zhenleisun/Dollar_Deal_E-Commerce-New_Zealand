<!--#include file="conn.asp"-->
<script language="javascript" src="../images/header.js">
</script>
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
<!--#include file="../md5.asp"-->
<%dim action,admin,adminid
action=request.QueryString("action")
adminid=request.QueryString("id")
admin=trim(request("admin"))
select case action
'//�޸�����
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from cnhww where adminid="&adminid,conn,1,3
rs("admin")=admin
if trim(request("password"))<>"" then
rs("password")=md5(trim(request("password")))
end if
rs("flag")=int(request("flag"))
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('�޸ĳɹ���');history.go(-1);</script>"
response.End
'//���������
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from cnhww",conn,1,3
rs.addnew
rs("admin")=trim(request("admin2"))
rs("password")=md5(trim(request("password2")))
rs("flag")=int(request("flag2"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('��ӳɹ���');history.go(-1);</script>"
response.End
'//ɾ������
case "del"
conn.execute ("delete from cnhww where adminid="&request.QueryString("id"))
response.Redirect "manageadmin.asp"
end select
%>
