
<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkkk()
{
     if(checkspace(document.myform.bookname.value)) {
	document.myform.bookname.focus();
    alert("��������Ʒ���ƣ�");
	return false;
  }
     if(checkspace(document.myform.shichangjia.value)) {
	document.myform.shichangjia.focus();
    alert("�������г��۸�");
	return false;
  }
     if(checkspace(document.myform.point.value)) {
	document.myform.point.focus();
    alert("������������֣�");
	return false;
  }
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
//-->
</script>
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
</head>
<body>
<form name="myform" method="post" action="SaveAward.asp?action=add" onSubmit="return checkkk()" >
  <table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
    <tr> 
      <th class="tableHeaderText" colspan=2>��ӽ�Ʒ</th>
    </tr>
    <tr> 
      <td width="30%" align="right" class="forumRowHighlight">��Ʒ���ƣ� </td>
      <td width="70%" class="forumRowHighlight"> 
        <input name="bookname" type="text" id="bookname" size="30">
      </td>
    </tr>
    <tr> 
      <td align="right" class="forumRowHighlight">�桡���� </td>
      <td class="forumRowHighlight"> <input name="guige" type="text" id="guige" size="20"> 
      </td>
    </tr>
    <tr> 
      <td align="right" class="forumRowHighlight">������֣� </td>
      <td class="forumRowHighlight">�ο��г��� 
        <input name="shichangjia" type="text" id="shichangjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="0">
        Ԫ������������� 
        <input name="point" type="text" id="point" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="50"> 
      </td>
    </tr>
    <tr> 
      <td rowspan="2" align="right" class="forumRowHighlight">��ƷͼƬ�� </td>
      <td class="forumRowHighlight"> <input name="bookpic" type="text" id="bookpic" size="30"> 
        &nbsp; <input class="button" type="button" name="Submit2" value="�ϴ�СͼƬ" onClick="window.open('../upload.asp?formname=myform&editname=bookpic&uppath=bookpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"> 
      </td>
    </tr>
    <tr> 
      <td class="forumRowHighlight"> <input name="bookpic2" type="text" id="bookpic2" size="30"> 
        &nbsp; <input class="button" type="button" name="Submit2" value="�ϴ���ͼƬ" onClick="window.open('../upload.asp?formname=myform&editname=bookpic2&uppath=bookpic&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"> 
      </td>
    </tr>
    <tr> 
      <td valign="top"  align="right" class="forumRowHighlight">��Ʒ˵���� </td>
      <td class="forumRowHighlight"> <textarea name="bookcontent" cols="46" rows="8" id="bookcontent"></textarea> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" class="forumRowHighlight"> <div align="center"> 
          <input name="xianshi" type="checkbox" id="xianshi" value="1">
          ��ʾ����ѡΪ���أ� ���� 
          <input class="button" type="submit" name="Submit" value="�� ��">
        </div></td>
    </tr>
  </table>
</form>
<br>
</body>
</html>