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
<html>
<head>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
    <form method="post" action="setunit.asp?action=update">
          <tr> 
            <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��λ����</font></b></td>
          </tr>
      <tr align=center > 
        <td width="20%" bgcolor="fbc2c2">�� ��</td>
        <td width="30%" bgcolor="fbc2c2">��λ����</td>
        <td width="30%" bgcolor="fbc2c2">�� ��(����������)</td>
        <td width="20%" bgcolor="fbc2c2">�� ��</td>
      </tr>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "SELECT  * from unit  order by danweiorder",conn,1,1
if rs.recordcount=0 then 
%>
<tr> 
<td colspan="8" height=25 align="center" width="100%" > ��û����ӵ�λ </td>
</tr>
<%
else
    do while not rs.eof
%>
      <tr > 
        <td align=center><%=rs("ID")%></td>
        <td align=center> 
          <input type="hidden" name="danweiid" value="<%=rs("id")%>">
          <input class=text type="text" name="danweiname" size="10" value="<%=trim(rs("danweiname"))%>">
        </td>
        <td align=center> 
          <input class=text type="text" name="danweiorder" size="10" value="<%=rs("danweiorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" >
        </td>
        <td align=center> 
          <a href="delunit.asp?danweiid=<%=rs("id")%>">ɾ��</a> 
        </td>
      </tr>
      <%
    rs.MoveNext
    Loop
%>
<tr> 
<td colspan="4"  align="center" width="100%"> 
<input type="submit" name="Submit2" value="�����޸�" style="font-family: ����; font-size: 9pt">&nbsp;
<input type="reset" value="�����趨" style="font-family: ����; font-size: 9pt" name="����" >
</td>
</tr>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</form>
</table><br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form method="post" action="setunit.asp?action=add">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��ӵ�λ</font></b></td>
</tr>
<tr height=40 > 
<td width=20% align=center>��ӵ�λ</td>
<td width=30% align=center> 
<input class=text type="text" name="danweiname" size="10">
</td>
<td width=30% align=center> 
<input class=text type="text" name="danweiorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td width=20% align=center> 
<input type="submit" name="Submit2" value="���" style="font-family: ����; font-size: 9pt">
</td>
</tr>
</form>
</table>
</body>
</html>
