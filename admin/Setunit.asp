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
dim danweiid,danweiorder,danweiname,danweino,i
if request("action")="update" then
    for i=1 to request.form("danweiid").count
	danweiid=replace(request.form("danweiid")(i),"'","")
	danweiorder=replace(request.form("danweiorder")(i),"'","")
	danweiname=replace(request.form("danweiname")(i),"'","")
	if replace(request.form("danweiorder")(i),"'","")="" then
%>
<script language=javascript>
history.back()
alert("����д��λ����ʾ����")
</script>
<%
	Response.End
	end if
	conn.execute("update unit set danweiorder="&danweiorder&",danweiname='"&danweiname&"' where id="&danweiid)
    next
conn.close
set conn=nothing
response.write "<script language=javascript>alert('��λ���óɹ���');window.location.href='manageunit.asp';</script>"
end if
if request("action")="add" then
	danweino=request.form("danweino")
	danweiname=trim(request("danweiname"))
	danweiorder=request.form("danweiorder")
	If danweiname="" Then
		response.write "��λ���Ʋ���Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
	If danweiorder="" Then
		response.write "������Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from unit",conn,1,3
rs.addnew
rs("danweiname")=danweiname
rs("danweiorder")=danweiorder
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script language=javascript>alert('��λ���ӳɹ���');window.location.href='manageunit.asp';</script>"
end if
%>