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
dim id
id=request("id")
if id<>"" then
if not isnumeric(id) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
	conn.execute("delete from zipcode where id=" &id)
	conn.close
	set conn=nothing
	response.write "<script language=javascript>alert('�ɹ�ɾ���ʱ࣡');window.location.href='managezipcode.asp';</script>"
	response.end
%>
