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
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
							  <tr>
							  <td colspan="3" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��Ʒ�鿴���޸�</font></b></td>
							  </tr>
                                <tr > 
                                  <td width="40%" align="center" bgcolor="fbc2c2">��������</td>
                                  <td width="20%" align="center" bgcolor="fbc2c2">��������</td>
                                  <td align="center" bgcolor="fbc2c2">ȷ������</td>
                                </tr>
                                <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * FROM bsort order by anclassidorder ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
		  %>
                                <form name="form1" method="post" action="savebsort.asp?action=edit&id=<%=int(rs("anclassid"))%>">
                                  <tr  align="center"> 
                                    <td><input name="anclass" type="text" id="anclass" size="12" value="<%=trim(rs("anclass"))%>">                                    </td>
                                    <td><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("anclassidorder"))%>">                                    </td>
                                    <td><input type="submit" name="Submit" value="�� ��">                                      &nbsp;
									  <a href="savebsort.asp?id=<%=int(rs("anclassid"))%>&action=del" onClick="return confirm('��ȷ��Ҫɾ���÷�����')"><font color=red>ɾ��</font></a>                                    </td>
                                  </tr>
                                </form>
						<%rs.MoveNext
								loop
								paixu=rs.RecordCount
								end if%>
</table>
<br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr>
<td colspan="3" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�����Ʒ����</font></b></td>
</tr>
<tr  align="center"> 
<td width="40%" align="center" bgcolor="fbc2c2"> ��������</td>
<td width="20%" align="center" bgcolor="fbc2c2"> ��������</td>
<td align="center" bgcolor="fbc2c2"> ȷ������</td>
</tr>
<form name="form1" method="post" action="savebsort.asp?action=add">
<tr  align="center"> 
<td>
<input name="anclass2" type="text" id="anclass2" size="12"></td>
<td>
<input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></td>
<td>
  <input type="submit" name="Submit3" value="�� ��"></td>
</tr>
</form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
