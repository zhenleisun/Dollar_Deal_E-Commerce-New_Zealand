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
dim pingpaiid
pingpaiid=request("pingpaiid")
	conn.execute("delete from brand where id=" &pingpaiid)
	conn.close
	set conn=nothing
	response.Write "<script language='javascript'>alert('ɾ���ɹ�!');window.location.href='ManageBrand.asp';</script>"
	response.End
%>