<html><head>
<title>����Ա��½</title>
<link href="../images/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
.MainHeader12 {
	PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-WEIGHT: bolder; FILTER: Shadow(Color=#c0c0c0,Direction=135); LEFT: 0px; PADDING-BOTTOM: 5px; LINE-HEIGHT: 11px; PADDING-TOP: 1px; HEIGHT: 5px; TEXT-ALIGN: left;
}
.netcst {
	position:absolute;
	top:0;
	right:0;
	height: 526px;
}
body {
	margin-top: 0px;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<% jianpankai=1 %>
<body leftmargin="0" marginwidth="0">
<table width="1003" height="480"  border="0" align="center" cellpadding="0" cellspacing="0" background="">
  <tr>
    <td height="584" align="center">
      <table width="98%" height="545"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="132">&nbsp;</td>
        </tr>
        <tr>
          <td align="center" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="364">
                          <tr>
                            <td width="364" height="103" valign="middle" align="center">
<form name="admininfo" method="post" action="chkadmin.asp" >
  <table width="355" border="0" align="center" cellpadding="2" cellspacing="1">
<tr>
  <td colspan="2"><div align="center"><b>��ע���¼����Ĵ�Сд</b></div></td>
  </tr>
<tr> 
<td width="38%"><div align="right">����Ա�ʺţ�
</div></td>
<td width="62%" align="center"><div align="left">
  <input class="wenbenkuang" name="admin" type="text" id="admin2" size="20"> 
</div></td>
</tr>
<tr> 
<td><div align="right">����Ա���룺
</div></td>
<td align="center"> 
  <div align="left"><input class="wenbenkuang" <% if jianpankai=1 then %>onclick= "password1=this;showkeyboard();this.readOnly=1;Calc.password.value=''"<% end if %> name="password" type="password" id="password" size="20"></div></td>
</tr>
<tr> 
<td><div align="right">������֤�룺</div></td>
<td align="center"><div align="left">
      <input class=wenbenkuang name=verifycode type=text value="<%If GetCode=9999 Then Response.Write "9999"%>" maxlength=4 size=10>  
      <a href='javascript:refreshimg()' title='�������������ͼƬ'><img  id='checkcode' src="../GetCode.asp" border="0" /></a></div></td>
</tr>
<tr>
  <td><div align="right">
    <input class="go-wenbenkuang" name="imageField" value="�� ¼" type="submit">    
  </div></td>
  <td align="center"><div align="left">
    <input class="go-wenbenkuang" type=reset name="Clear" value="�� ��">
                            <input class="go-wenbenkuang" type=button onClick="window.open('../index.asp')" name="Clear2" value="������ҳ">
                          </div></td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td align="center">&nbsp;</td>
</tr>
</table>
</form>							</td>
                          </tr>
                      </table></td>
        </tr>
        <tr>
          <td height="108">&nbsp;</td>
        </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<% if jianpankai=1 then %><script language="JavaScript" src="images/softkeyboard.js"></script>
<%end if %>
    <script type="text/javascript">
    function refreshimg(){document.all.checkcode.src='../getcode.asp';}
    </script>