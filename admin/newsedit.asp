<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
dim newsid
newsid=request.QueryString("id")
if not isnumeric(newsid) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
if request.QueryString("action")="save" then 
Content=Request.Form("Content")

set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news where newsid="&newsid,conn,1,3
rs("newsname")=trim(request("newsname"))
rs("addname")=trim(request("addname"))
rs("newscontent")=trim(request("content"))
' rs("adddate")=now()
rs.update
rs.close
set rs=nothing
session("content")=""
response.write "<script language=javascript>alert('更改成功！');window.location.href='"&request("linkaddress")&"';</script>"
response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.newsname.value=="")
	{
	  alert("对不起，请输入新闻主题！")
	  document.form1.newsname.focus()
	  return false
	 }
if (document.form1.viewhtml.checked == true)
	{
	  alert("对不起，请取消“查看HTML源代码”后再添加！")
	  document.form1.viewhtml.focus()
	  return false
	 }
if (document.form1.Content.value.length==0)
	{
	  alert("对不起，请输入文章内容！")
	  //document.form1.content.focus()
	  return false
	 }
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<script  language="JavaScript">
<!-- Hide from older browsers...
//Function to format text in the text box
function FormatText(command, option){
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}
//Function to add image
function AddImage(){	
	imagePath = prompt('请输入图片地址', 'http://');				
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
	frames.message.focus();			
}
//Function to clear form
function ResetForm(){
	if (window.confirm('确认要清空对话框内容?')){
	 	frames.message.document.body.innerHTML = ''; 
	 	return true;
	 } 
	 return false;		
}
//Function to open pop up window
function openWin(theURL,winName,features) {
  	window.open(theURL,winName,features);
}
function setMode(newMode)
{
  bTextMode = newMode;
  var cont;
  if (bTextMode) {
    cleanHtml();
    cleanHtml();
    cont=message.document.body.innerHTML;
    message.document.body.innerText=cont;
  } else {
    cont=message.document.body.innerText;
    message.document.body.innerHTML=cont;
  }
  message.focus();
}
function cleanHtml()
{
  var fonts = message.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }
}
// -->
</script>
<script language="JavaScript">
function help()
{
    var helpmess;
    helpmess="---------------填写帮助---------------\r\n\r\n"+
         "1.请不要发表有危险性的脚本。\r\n\r\n"+
         "2.如果要书写源代码，请选中\r\n\r\n"+
         "　查看HTML源代码书写.\r\n\r\n"+
         "3.需要你自己运行,才能看效果.\r\n\r\n"+
         "4.如果书写js，尽量不要在这儿书写.\r\n\r\n";
    alert(helpmess);
}
</script>
<script src="edit.js" type="text/javascript"></script>
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<%

Dim htmlData

htmlData = Request.Form("newscontent")

Function htmlspecialchars(str)
	str = Replace(str, "&", "&amp;")
	str = Replace(str, "<", "&lt;")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, """", "&quot;")
	htmlspecialchars = str
End Function
%>
<script type="text/javascript" charset="utf-8" src="../kindeditor.js"></script>
	<script type="text/javascript">
		KE.show({
			id : 'newscontent',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>
<body>

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news where newsid="&newsid,conn,1,1%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td height="25" colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">更改网站新闻</font></b></td>
</tr>
<tr> 
<td valign="top"> 
<form name="form1" method="post" action="newsedit.asp?action=save&id=<%=newsid%>" OnSubmit="return checkdata()" onReset="return ResetForm();">
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
          <tr > 
            <td width="30%" align="right" bgcolor="fbf4f4">新闻主题：</td>
            <td width="70%" bgcolor="fbf4f4" style="PADDING-LEFT: 10px"> <input name="newsname" type="text" id="newsname" value="<%=rs("newsname")%>"> 
            </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="fbf4f4">发 表 人：</td>
            <td bgcolor="fbf4f4" style="PADDING-LEFT: 10px"> <input name="addname" type="text" id="addname" value="<%=trim(rs("addname"))%>"> 
            </td>
          </tr>
          <tr > 
            <td align="right" valign="top" bgcolor="fbf4f4">新闻内容：</td>
            <td class="forumRowHighlight" style="PADDING-LEFT: 10px"><textarea id="newscontent" name="content" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"><%=htmlspecialchars(rs("newscontent"))%></textarea>
              &nbsp;</td>
          </tr>
          <tr> 
            <td align="right" bgcolor="fbf4f4" ></td>
            <td height="30" bgcolor="fbf4f4"  style="PADDING-LEFT: 10px"> <input type="submit" name="Submit" value="修改保存" OnClick="document.form1.Content.value = frames.message.document.body.innerHTML;"> 
              <input type="button" value=" 返 回 " onClick="javascript:history.go(-1)" class="unnamed5" name="button"> 
              <input type="hidden" name="Content" value=""> <input type="hidden" name="linkaddress" value="<%=request.servervariables("http_referer")%>"> 
            </td>
          </tr>
        </table>
</form>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
