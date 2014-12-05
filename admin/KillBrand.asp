<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<%
dim pingpaiid
pingpaiid=request("pingpaiid")
	conn.execute("delete from brand where id=" &pingpaiid)
	conn.close
	set conn=nothing
	response.Write "<script language='javascript'>alert('删除成功!');window.location.href='ManageBrand.asp';</script>"
	response.End
%>