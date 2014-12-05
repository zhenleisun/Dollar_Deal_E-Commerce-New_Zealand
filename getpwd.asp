<html>
<head>
<title>密码找回</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="javascript">
<!--
function form1_onsubmit() {
if (document.form1.username.value=="")
	{
	alert("Please fill in the user name！")
	document.form1.username.focus()
	return false
	}
}
// --></script>
</head>
<body topmargin="20" leftmargin="0" oncontextmenu="self.event.returnValue=false">
<form method="POST" name="form1" language="javascript" onSubmit="return form1_onsubmit()" action="getpwd2.asp">
<table border="0" cellpadding="2" cellspacing="1" width="400" bgcolor="#6FABCB" style="border-collapse: collapse" align="center">
<TR><TD bgColor=#ffffff height=1></TD></TR>
<TR><TD align=center>The first step: fill in your name.</TD>
</TR>
<tr bgcolor="#ffffff"> 
<td height="100" bgcolor="F8FAFB"> 
<table border="0" cellpadding="0" width="100%" cellspacing="0">
<tr> 
<td width="100%" height="22" align="center">Forgot your password? Please add your name here！</td>
</tr>
<tr> 
<td width="100%" align="center">
<input class="wenbenkuang" type="text" name="username" size="18">
<input class="go-wenbenkuang" name="imageField" type="submit" value=" Next " onFocus="this.blur()">
</td>
</tr>
</table>
<p align="center"><a href="javascript:window.close()" class="1">Close</a></p>
</td>
</tr>
</table>
</form>
</body>
</html>