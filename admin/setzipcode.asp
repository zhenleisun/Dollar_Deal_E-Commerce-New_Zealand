<!--#include file="conn.asp"-->
<link href="../images/css.css" rel="stylesheet" type="text/css">
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
dim id,youbian,dizhiname,danweino,i
if request("action")="update" then
    for i=1 to request.form("id").count
	id=replace(request.form("id")(i),"'","")
	youbian=replace(request.form("youbian")(i),"'","")
	dizhiname=replace(request.form("dizhiname")(i),"'","")
	if len(youbian)<>6 then
%>
<script language=javascript>
history.back()
alert("����ȷ��д�ʱ࣡")
</script>
<%
		Response.End
	end if
	conn.execute("update zipcode set youbian="&youbian&",dizhiname='"&dizhiname&"' where id="&id)
    next
conn.close
set conn=nothing
response.write "<script language=javascript>alert('�ʱ����óɹ���');window.location.href='managezipcode.asp';</script>"
end if
if request("action")="add" then
	dizhiname=trim(request("dizhiname"))
	youbian=request.form("youbian")
	If len(youbian)<>6 Then
		response.write "<script language=javascript>alert('����ȷ��д�ʱ࣡�뷵��������д��');history.go(-1);</script>"
		response.end
	end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from zipcode",conn,1,3
rs.addnew
rs("dizhiname")=dizhiname
rs("youbian")=youbian
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "<script language=javascript>alert('�ʱ���ӳɹ���');window.location.href='managezipcode.asp';</script>"
end if
%>
