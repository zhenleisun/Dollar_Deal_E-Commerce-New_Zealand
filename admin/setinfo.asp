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
<%dim action
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.Open "select webname,musicd,musicurl,webemail,dizhi,xp,dhxg,tjj,tejia,youbian,gdong,nocopy,daohang,dianhua,copyright,weblogo,weburl,webbanner,icp,qq,mailaddress,mailsend,mailusername,mailname,mailuserpass,view from webinfo ",conn,1,3
rs("webname")=trim(request("webname"))
rs("webemail")=trim(request("webemail"))
rs("dizhi")=trim(request("dizhi"))
rs("youbian")=trim(request("youbian"))
rs("dianhua")=trim(request("dianhua"))
rs("copyright")=trim(request("copyright"))
rs("webbanner")=trim(request("webbanner"))
rs("nocopy")=trim(request("nocopy"))
rs("daohang")=trim(request("daohang"))
rs("weblogo")=trim(request("weblogo"))
rs("gdong")=trim(request("gdong"))
rs("webbanner")=trim(request("webbanner"))
rs("weburl")=trim(request("weburl"))
rs("icp")=trim(request("icp"))
rs("qq")=trim(request("qq"))
rs("dhxg")=trim(request("dhxg"))
rs("xp")=trim(request("xp"))
rs("tjj")=trim(request("tjj"))
rs("tejia")=trim(request("tejia"))
rs("view")=trim(request("view"))
rs("mailaddress")=trim(request("mailaddress"))
rs("mailsend")=trim(request("mailsend"))
rs("mailusername")=trim(request("mailusername"))
rs("mailuserpass")=trim(request("mailuserpass"))
rs("mailname")=trim(request("mailname"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('��վ�����޸ĳɹ���');history.go(-1);</script>"
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
  <form name="form1" method="post" action="setinfo.asp?action=save">
    <%set rs=server.CreateObject("adodb.recordset")
rs.Open"select webname,musicurl,musicd,webemail,dizhi,youbian,xp,dhxg,tjj,nocopy,daohang,tejia,dianhua,gdong,copyright,weblogo,weburl,webbanner,icp,qq,mailaddress,mailsend,mailusername,mailname,mailuserpass,view from webinfo",conn,1,1

%>
    <tr> 
      <td colspan="2" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��վ��Ϣ����</font></b></td>
    </tr>
    <tr > 
      <td width="30%" align="right">��վ��ַ��</td>
      <td width="70%" style="PADDING-LEFT: 10px">HTTP:// 
        <input name="weburl" type="text" id="weburl" size="21" value="<%=trim(rs("weburl"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��վ���ƣ�</td>
      <td style="PADDING-LEFT: 10px"> <input name="webname" type="text" id="webname" size="28" value="<%=trim(rs("webname"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��˾ e-mail��</td>
      <td style="PADDING-LEFT: 10px"> <input name="webemail" type="text" id="webemail" size="28" value="<%=trim(rs("webemail"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��˾��ַ��</td>
      <td style="PADDING-LEFT: 10px"> <input name="dizhi" type="text" id="dizhi" size="28" value="<%=trim(rs("dizhi"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��˾�ʱࣺ</td>
      <td style="PADDING-LEFT: 10px"> <input name="youbian" type="text" id="youbian" size="28" value="<%=trim(rs("youbian"))%>"	onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��˾�绰��</td>
      <td style="PADDING-LEFT: 10px"> <input name="dianhua" type="text" id="dianhua" size="28" value="<%=trim(rs("dianhua"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td align="right">��վ��Ȩ��</td>
      <td style="PADDING-LEFT: 10px"> <input name="copyright" type="text" size="28" value="<%=trim(rs("copyright"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr> 
      <td align="right">������ѯQQ(��Ѷ)��</td>
      <td style="PADDING-LEFT: 10px"> <input name="webbanner" type="text" id="webbanner" size="28" value="<%=trim(rs("webbanner"))%>"> 
        <font color=#ff0000>**</font>������QQ����Ӣ�Ķ��Ÿ���</td>
    </tr>
    <tr> 
      <td align="right">�ͷ�QQ״��</td>
      <td style="PADDING-LEFT: 10px"> <input name="qq" type="radio" value="1"   <% if rs("qq")=1 then %> checked <%end if %>>
        �� 
        <input type="radio" name="qq"  <% if rs("qq")=2 then %> checked <%end if %> value="2">
        �ر�</td>
    </tr>
    <tr> 
      <td align="right">������˹��ܣ�</td>
      <td style="PADDING-LEFT: 10px"> <input name="view" type="radio" value="0"   <% if rs("view")=0 then %> checked <%end if %>>
        �� 
        <input type="radio" name="view"  <% if rs("view")=1 then %> checked <%end if %> value="1">
        �ر�</td>
    </tr>
    <tr> 
      <td align="right">�Ƿ���ʾ��ҳ�����Ƽ���</td>
      <td style="PADDING-LEFT: 10px"><input name="gdong" type="radio" value="1"   <% if rs("gdong")=1 then %> checked <%end if %>>
        �� 
        <input type="radio" name="gdong"  <% if rs("gdong")=2 then %> checked <%end if %> value="2">
        �ر�</td>
    </tr>
    <tr>
      <td align="right">���������Ƿ���ʾƷ�Ʒ��ࣺ</td>
      <td style="PADDING-LEFT: 10px"><input type="radio" name="daohang" value="0"  <% if rs("daohang")=0 then %> checked <%end if %>>
        ��
        <input type="radio" name="daohang" value="1" <% if rs("daohang")=1 then %> checked <%end if %>>
        �ر� </td>
    </tr>
    <tr>
      <td align="right">��ֹ������ҳ���ϣ�</td>
      <td style="PADDING-LEFT: 10px"><input type="radio" name="nocopy" value="0"  <% if rs("nocopy")=0 then %> checked <%end if %>>
        ��
        <input type="radio" name="nocopy" value="1" <% if rs("nocopy")=1 then %> checked <%end if %>>
        �ر� </td>
    </tr>
    <tr> 
      <td align="right">�������˵�Ч����</td>
      <td style="PADDING-LEFT: 10px"><input name="dhxg" type="radio" value="0"   <% if rs("dhxg")=0 then %> checked <%end if %>>
        ����ʽ &nbsp <input type="radio" name="dhxg"  <% if rs("dhxg")=1 then %> checked <%end if %> value="1">
        ͼƬʽ &nbsp </td>
    </tr>
    <tr> 
      <td align="right"> ��ҳ��Ʒ��ʾ������</td>
      <td style="PADDING-LEFT: 10px">��Ʒ 
        <input name="xp" type="text" id="xinpin" value="<%=trim(rs("xp"))%>" size="5">
        �����Ƽ� 
        <input name="tjj" type="text" id="tuijian" value="<%=trim(rs("tjj"))%>" size="5">
        �����ؼ� 
        <input name="tejia" type="text" id="xp3" value="<%=trim(rs("tejia"))%>" size="5">
        ��</td>
    </tr>
    <tr > 
      <td align="right">Icp�����ţ�</td>
      <td style="PADDING-LEFT: 10px"> <input name="icp" type="text" size="28" value="<%=trim(rs("icp"))%>"> 
        <font color=#ff0000>**</font></td>
    </tr>
    <tr > 
      <td height="-1" align="right">��վLogo��</td>
      <td style="PADDING-LEFT: 10px"> <input name="weblogo" type="text" id="weblogo" size="28" value="<%=trim(rs("weblogo"))%>"> 
        <font color=#ff0000>**</font> <input type="button" name="Submit2215" value="�ϴ�ͼƬ" onClick="window.open('../upload.asp?formname=form1&editname=weblogo&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
    </tr>
    <tr  > 
      <td colspan="2" align="center">�����ʼ������� </td>
    </tr>
    <tr> </tr>
    <tr  > 
      <td align="right">�ʼ���������ַ��</td>
      <td style="PADDING-LEFT: 10px"> <input name="mailaddress" type="text" id="mailaddress" size="36" value="<%=trim(rs("mailaddress"))%>">      </td>
    </tr>
    <tr  > 
      <td align="right">�������䣺</td>
      <td style="PADDING-LEFT: 10px"> <input name="mailsend" type="text" id="mailsend" size="36" value="<%=trim(rs("mailsend"))%>">      </td>
    </tr>
    <tr  > 
      <td align="right">�� ¼ ����</td>
      <td style="PADDING-LEFT: 10px"> <input name="mailusername" type="text" id="mailusername" size="36" value="<%=trim(rs("mailusername"))%>">      </td>
    </tr>
    <tr  > 
      <td align="right">��¼���룺</td>
      <td style="PADDING-LEFT: 10px"> <input name="mailuserpass" type="password" id="mailuserpass" size="36" value="<%=trim(rs("mailuserpass"))%>">      </td>
    </tr>
    <tr  > 
      <td align="right">��ʾ��������</td>
      <td style="PADDING-LEFT: 10px"> <input name="mailname" type="text" id="mailname" size="36" value="<%=trim(rs("mailname"))%>">      </td>
    </tr>
    <tr > 
      <td></td>
      <td style="PADDING-LEFT: 10px"> <input type="submit" name="Submit" value=" �޸ı��� "> 
        &nbsp; <input type="reset" name="Submit2" value=" ������д "></td>
    </tr>
  </form>
</table>
<%rs.Close
  set rs=nothing
  %>
<!--#include file="foot.asp"-->
</body>
</html>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>
