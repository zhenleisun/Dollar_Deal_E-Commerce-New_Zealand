<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
dim i
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../images/css.css" rel=stylesheet>
<SCRIPT language="JavaScript" type="text/javascript">
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<form method="POST" action="Savevote.asp">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�����µ�ͶƱ</font></b></td>
</tr>
<tr > 
<td width="20%" align="right">ͶƱ���⣺</td>
<td width="80%"> 
<input type="text" name="Title" size="40" style="font-family: ����; font-size: 9pt"></td>
</tr>
<%for i=1 to 8%>
<tr > 
<td align="right">ѡ �� <%=i%>��</td>
<td> 
<input type="text" name="select<%=i%>" size="25" style="font-family: ����; font-size: 9pt">
Ʊ ����<input type="text" name="answer<%=i%>" size="5" value="0" style="font-family: ����; font-size: 9pt" onKeyPress="event.returnValue=IsDigit();" ></td>
</tr>
<%next%>
<tr >
<td></td>
<td>
<input type="hidden" value="add" name="act">
<input type="submit" value=" �� �� " name="cmdok" style="font-family: ����; font-size: 9pt">&nbsp;
<input type="reset" value=" ȡ �� "  name="cmdcancel" style="font-family: ����; font-size: 9pt">
<input type="button" value=" �� �� " onClick="javascript:history.go(-1)" style="font-family: ����; font-size: 9pt">&nbsp;
</td>
</tr>
<div align="center"></div>
</form>
</table>
</body></html>