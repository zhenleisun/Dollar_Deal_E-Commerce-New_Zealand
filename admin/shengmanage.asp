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
<html><head><title>ʡ����</title>
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
<form method="post" action="shengset.asp?action=update">
<tr> 
<td class="forumRowHighlight" colspan="5" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">ʡ�����޸�</font></b></td>
</tr>
<tr align=center> 
<td class="forumRowHighlight" width="10%">���</td>
<td class="forumRowHighlight" width="30%">ʡ����</td>
<td class="forumRowHighlight" width="30%">ʡ���</td>
<td class="forumRowHighlight" width="20%">����<br>(���������)</td>
<td class="forumRowHighlight" width="10%">����</td>
</tr>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "SELECT  * From province  order by shengorder",conn,1,1
if rs.recordcount=0 then 
%>
<tr align=center>
<td class="forumRowHighlight" colspan="5" align="center" width="100%">��û������ʡ</td>
</tr>
<%
	else
    do while not rs.eof
%>
<tr align=center> 
<td class="forumRowHighlight" align=center><%=rs("ID")%></td>
<td class="forumRowHighlight" align=center> 
<input type="hidden" name="shengid" value="<%=rs("id")%>">
<input class=text type="text" name="shengname" size="10" value="<%=trim(rs("shengname"))%>">
</td>
<td class="forumRowHighlight" align=center> 
<input class=text type="text" name="shengno" size="10" value="<%=trim(rs("shengno"))%>" ONKEYPRESS="event.returnValue=IsDigit();" >
</td>
<td class="forumRowHighlight" align=center> 
<input class=text type="text" name="shengorder" size="10" value="<%=rs("shengorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" >
</td>
<td class="forumRowHighlight" align=center> 
<a href="shengkill.asp?shengid=<%=rs("id")%>">ɾ��</a> 
</td>
</tr>
	<%
	rs.MoveNext
	Loop
%>
<tr align=center> 
<td class="forumRowHighlight" colspan="8" height=25 align="center" width="100%"> 
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
<tr> 
<td class="forumRowHighlight" colspan="5" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��Ʒ��Ѷ</font></b></td>
</tr>
<form method="post" action="shengset.asp?action=add">
<tr align=center> 
<td class="forumRowHighlight" width=10% align=center>����ʡ 
</td>
<td class="forumRowHighlight" width=30% align=center> 
<input class=text type="text" name="shengname" size="10">
</td>
<td class="forumRowHighlight" width=30% align=center> 
<input class=text type="text" name="shengno" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td class="forumRowHighlight" width=20% align=center> 
<input class=text type="text" name="shengorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
</td>
<td class="forumRowHighlight" width=10% align=center> 
<input type="submit" name="Submit2" value="����" style="font-family: ����; font-size: 9pt">
</td>
</tr>
</form>
</table>
</body>
</html>