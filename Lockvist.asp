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
<font size="2">提示：您的IP地址<font color=red> <%=Request.serverVariables("REMOTE_ADDR")%></font> 被限制访问，请与管理员联系。</font>
	</td></tr>
	</table>
<%
response.end
end if
next
end if
end if
%>