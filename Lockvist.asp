<%

if webnow=1 then 
%><br><br>
<table width="560" height="290" background="images/error.gif" cellspacing="0" cellpadding="0" align="center">
<tr><td align="center"  valign="middle">
<font size="2"><%=webmess%></font>
	</td></tr>
	</table>
<%
response.end 
end if 


if lockip<>"0" then
if ip<>"" then
lockip=split(cstr(ip),"@")
for N=0 to UBound(lockip)
if instr(Request.serverVariables("REMOTE_ADDR"),lockip(n))>0 then
%><br><br>
<table width="560" height="290" background="images/error.gif" cellspacing="0" cellpadding="0" align="center">
<tr><td align="center"  valign="middle">
<font size="2">��ʾ������IP��ַ<font color=red> <%=Request.serverVariables("REMOTE_ADDR")%></font> �����Ʒ��ʣ��������Ա��ϵ��</font>
	</td></tr>
	</table>
<%
response.end
end if
next
end if
end if
%>