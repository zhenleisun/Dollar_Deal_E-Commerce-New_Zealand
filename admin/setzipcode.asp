<!--#include file="conn.asp"-->
<link href="../images/css.css" rel="stylesheet" type="text/css">
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
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
alert("请正确填写邮编！")
</script>
<%
		Response.End
	end if
	conn.execute("update zipcode set youbian="&youbian&",dizhiname='"&dizhiname&"' where id="&id)
    next
conn.close
set conn=nothing
response.write "<script language=javascript>alert('邮编设置成功！');window.location.href='managezipcode.asp';</script>"
end if
if request("action")="add" then
	dizhiname=trim(request("dizhiname"))
	youbian=request.form("youbian")
	If len(youbian)<>6 Then
		response.write "<script language=javascript>alert('请正确填写邮编！请返回重新填写！');history.go(-1);</script>"
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
response.write "<script language=javascript>alert('邮编添加成功！');window.location.href='managezipcode.asp';</script>"
end if
%>
