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

dim shengid
shengid=request("shengid")
if shengid<>"" then
if not isnumeric(shengid) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
	conn.execute("delete from city where shengid=" &shengid)
	conn.execute("delete from province where id=" &shengid)
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('��ѡ���ʡ�Ѿ���ɾ����');window.location.href='shengmanage.asp';</script>"
	response.end
%>