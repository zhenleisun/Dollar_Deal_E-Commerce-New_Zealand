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
<%dim anclassid,anclass,paixu
anclass=request.QueryString("nclass")
anclassid=request.QueryString("id")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��ƷС�����</font></b></td>
</tr>
<tr> 
<td width="30%" valign="top" align="center" ><br>
<select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
<option >ѡ����Ʒ����</option>
                                  <%set rs=server.createobject("adodb.recordset")
		rs.Open "select * from bsort order by anclassidorder",conn,1,1
		do while not rs.eof %>
                                  <option value="managessort.asp?id=<%=rs("anclassid")%>&anclass=<%=rs("anclass")%>" <%if rs("anclassid")=cint(request.QueryString("id")) then %>selected<%end if%>><%=trim(rs("anclass"))%></option>
                                  <%rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
                                </select><br>
                                <%if request.QueryString("id")<>"" then
        response.Write "��ǰ��Ѷ��"&request.QueryString("anclass")
        end if%>
</td>
<td width="70%" > 
<table width="90%" align="center" border="0" cellpadding="1" cellspacing="2">
<tr> 
<td width="50%" align="center">��������</td>
<td width="20%" align="center">��������</td>
<td width="30%" align="center">ȷ������</td>
</tr>
                                <%
        if anclassid="" then
        response.Write "<div align=center><font color=red>��ѡ�����ķ���</font></div>"
        else
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * FROM ssort where anclassid="&anclassid&" order by nclassidorder",conn,1,1
         if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  paixu=0
		  else
         do while not rs.EOF
         %>
<form name="form1" method="post" action="savessort.asp?action=edit&id=<%=rs("nclassid")%>&anclass=<%=request.QueryString("anclass")%>">
<tr> 
<td align="center">
<input name="nclass" type="text" id="nclass" size="16" value="<%=trim(rs("nclass"))%>">
<input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>" id="Hidden1">
</td>
<td align="center">
<input name="nclassidorder" type="text" id="nclassidorder" size="4" value="<%=int(rs("nclassidorder"))%>">
</td>
<td align="center">
<input type="submit" name="Submit" value="�� ��">&nbsp;
<a href="savessort.asp?id=<%=int(rs("nclassid"))%>&action=del&anclassid=<%=request.QueryString("id")%>&anclass=<%=request.QueryString("anclass")%>" onClick="return confirm('��ȷ������ɾ��������')"><font color=red>ɾ��</font></a> 
</td>
</tr>
</form>
		<%rs.movenext
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        end if
		%>
</table>
</td>
</tr>
</table><br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�����Ʒ����</font></b></td>
</tr>
<tr> 
<td width="30%" valign="top" align="center" ><br>
��ǰ���ࣺ<%=request.QueryString("anclass")%></div>
</td>
<td width="70%" > 
<table width="90%" align="center" border="0" cellpadding="1" cellspacing="2">
<tr> 
<td width="50%" align="center">��������</div></td>
<td width="20%" align="center">��������</div></td>
<td width="30%" align="center">ȷ������</div></td>
</tr>
<form name="form2" method="post" action="savessort.asp?action=add&anclass=<%=request.QueryString("anclass")%>">
<tr> 
<td align="center"> 
<input name="nclass2" type="text" id="nclass22" size="16">
<input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>">
</td>
<td align="center"> 
<input name="nclassidorder2" type="text" id="nclassidorder22" size="4" value="<%=paixu+1%>">
</td>
<td align="center"> 
<input type="submit" name="Submit2" value="��ӷ���">
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
