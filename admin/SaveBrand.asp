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
dim pingpaiid,pingpaiorder,pingpainame,tuijian,i
if request("action")="update" then
    for i=1 to request.form("pingpaiid").count
	pingpaiid=replace(request.form("pingpaiid")(i),"'","")
	pingpaiorder=replace(request.form("pingpaiorder")(i),"'","")
	tuijian=replace(request.form("tuijian")(i),"'","")
	pingpainame=replace(request.form("pingpainame")(i),"'","")
	if replace(request.form("pingpaiorder")(i),"'","")="" then
%>
<script language=javascript>
history.back()
alert("����дƷ�Ƶ���ʾ����")
</script>
<%
Response.End
end if
conn.execute("update brand set pingpaiorder="&pingpaiorder&",pingpainame='"&pingpainame&"',tuijian="&tuijian&" where id="&pingpaiid)
next
conn.close
set conn=nothing
response.Write "<script language='javascript'>alert('Ʒ�����óɹ�!��');window.location.href='ManageBrand.asp';</script>"
response.End
end if

if request("action")="add" then
	tuijian=request.form("tuijian")
	pingpainame=trim(request("pingpainame"))
	pingpaiorder=request.form("pingpaiorder")
	If pingpainame="" Then
		response.write "Ʒ�����Ʋ���Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
	If pingpaiorder="" Then
		response.write "������Ϊ�գ���<a href=javascript:history.go(-1)>����������д</a>��"
		response.end
	end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "select * from brand",conn,1,3
rs.addnew
rs("tuijian")=tuijian
rs("pingpainame")=pingpainame
rs("pingpaiorder")=pingpaiorder
rs.update
rs.close
set rs=nothing

conn.close
set conn=nothing
response.Write "<script language='javascript'>alert('Ʒ�����ӳɹ�!��');window.location.href='ManageBrand.asp';</script>"
response.End
end if
%>