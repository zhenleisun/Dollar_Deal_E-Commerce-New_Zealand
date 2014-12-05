<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="css.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value+"*"+j.value
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
<form name=flash>
<iframe name="ad" frameborder=0 width=100% height=20 scrolling=no src="../upme1.htm"></iframe><br>
&nbsp;&nbsp;动画类型：swf，上传限制：2000K。<br>&nbsp;&nbsp;动画地址：<input id=a name=a value='' style='width:200' title='可直接输入网上Flash动画的地址，或由上传程序自动产生Flash动画地址' style="border: 1px solid #0000FF" size="20"><br>
&nbsp;&nbsp;动画注释：<input id=h name=h style='width:200' title='该注释将自动显示在动画的下方' style="border: 1px solid #0000FF" size="20"><input id=j type="hidden" name=j value='' style='width:200' title='可直接输入网上Flash动画的地址，或由上传程序自动产生Flash动画地址' style="border: 1px solid #0000FF" size="20"><br>
&nbsp;&nbsp;立即播放：<select id=b><option value="-1">是<option value="0">否</select>
&nbsp;&nbsp;循环播放：<select id=c><option value="-1">是<option value="0">否</select><br>
&nbsp;&nbsp;显示菜单：<select id=d ><option value="-1">是<option value="0">否</select>
&nbsp;&nbsp;文件位置：<select id=g ><option value="left">居左<option value="center">居中<option value="right">居右</select><br>
&nbsp;&nbsp;<font color=red>立即播放和显示菜单不要同时选“否”</font><br>
&nbsp;&nbsp;<FONT 
color=red>文件位置功能能实现图文环绕，但选择后将不能更改</FONT><BR>&nbsp;&nbsp;动画高度：<input id=e value='500' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" style="border: 1px solid #0000FF" size="20">&nbsp;&nbsp;动画宽度：<input id=f value='375' style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" style="border: 1px solid #0000FF" size="20">
<br>&nbsp;&nbsp;<input type=button value='插入' id=Ok title='点击“插入”按钮，在编辑器中插入该Flash动画' style="border: 1px solid #0000FF"></form>