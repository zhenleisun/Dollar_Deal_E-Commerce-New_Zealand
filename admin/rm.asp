<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="css.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value
window.close();
</SCRIPT>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</HEAD>
<BODY>
<form name=rm>
&nbsp;&nbsp;�ļ���ַ��<input id=a name=a value='http://' style='width:200' title='��ֱ����������RealPlay�ļ��ĵ�ַ�������ϴ������Զ�����RealPlay�ļ��ĵ�ַ' size="20"><br>
&nbsp;&nbsp;��������RealPlay�ļ�֧�ָ�ʽΪ��rm��ra��ram<br>
&nbsp;&nbsp;�ļ�ע�ͣ�<input id=g name=g style='width:200' title='��ע�ͽ��Զ���ʾ���ļ����·�' size="20"><br>
&nbsp;&nbsp;�������ţ�<select id=b><option value="-1">��<option value="0">��</select>
&nbsp;&nbsp;ѭ�����ţ�<select id=c><option value="-1">��<option value="0">��</select><br>
&nbsp;&nbsp;�ļ�λ�ã�<select id=d><option value="left">����<option value="center">����<option value="right">����</select>
&nbsp;&nbsp;��ʾѡ��<select id=h><option value="0">ֻ������<option value="1">��ʾ��Ƶ</select><br>
&nbsp;&nbsp;���Ÿ߶ȣ�<input id=e value='352' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" size="20">&nbsp;&nbsp;���ſ�ȣ�<input id=f value='240' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" size="20">
<br>&nbsp;&nbsp;<input type=button value='����' id=Ok title='��������롱��ť���ڱ༭���в����RealPlay�ļ�'></form>