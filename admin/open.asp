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
%>
<script language="JavaScript" type="text/JavaScript">
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
  
  }
  else
  {
   linkimg.style.display="";
  
  }
}	
</script>
<%dim action
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from webinfo ",conn,1,3
rs("webnow")=trim(request("webnow"))

rs("webmess")=trim(request("webmess"))

rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('网站资料修改成功！');history.go(-1);</script>"
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
  <form name="form1" method="post" action="open.asp?action=save">
    <%set rs=server.CreateObject("adodb.recordset")
rs.Open"select * from webinfo",conn,1,1

%>
    <tr> 
      <td colspan="2" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">网站开关设置</font></b></td>
    </tr>
    <tr class="tableborder"> 
      <td width="30%" align="right">网站开关：</td>
      <td width="70%"> <input type="radio" value="0" name="webnow" <%if rs("webnow")=0 then response.write "checked" %> onClick='showlist("a");'>
        开启&nbsp;&nbsp; <input type="radio" value="1" name="webnow" <%if rs("webnow")=1 then response.write "checked" %> onClick='showlist("b");'>
        关闭</td>
    </tr>
    <tr id="linkimg" <%if root_info_OnOff=0 then%>style='display:none'<%end if%>> 
      <td align="right" >关闭提示语：</td>
      <td><textarea name="webmess" cols="47" rows="4" id="webmess"><%=rs("webmess")%></textarea></td>
    </tr>
    <tr > 
      <td></td>
      <td style="PADDING-LEFT: 10px"> <input type="submit" name="Submit" value=" 修改保存 "> 
        &nbsp; <input type="reset" name="Submit2" value=" 重新添写 "></td>
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



