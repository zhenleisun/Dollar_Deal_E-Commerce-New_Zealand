<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<%
NewsID=request("id")
if NewsID<>"" then
if not isnumeric(NewsID) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
 
set rs=server.CreateObject ("ADODB.RecordSet")
rs.Source="select * FROM inform where NewsID=" & NewsID
rs.Open rs.source,conn,1,1
if rs.recordcount=0 then
%>
<script language=javascript>
history.back()
alert("非法文章ID，请确认NewsID不超出当前范围！")
</script>
<%
Response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
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
 
 
</head>
<body>

<form name="form1" method="post" action="editinform.asp" OnSubmit="return checkdata()" onReset="return ResetForm();">
  <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
    <tr> 
      <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">修改专题文章-该文章ID为：<%=NewsID%></font></b></td>
    </tr>
    <tr > 
      <td width="20%" align="right" class="TableBody1" valign="top">文章标题：</td>
      <td width="80%" class="TableBody1"> <input name="title" id=me type="TEXT" size=50 maxlength=100 style="background-color:ffffff;border: 1 double"  onMouseOver="window.status='在这里输入要添加的文章标题，必填';return true;" onMouseOut="window.status='';return true;" title='在这里输入要添加的文章标题，必填' value="<%=rs("title")%>"> 
        <span style='cursor:hand' title='加长对话框' onClick='if (me.size<102)me.size=me.size+2'>+</span> 
        <span style='cursor:hand' title='缩短对话框' onClick='if (me.size>10)me.size=me.size-2'>-</span> 
        <br>
      </td>
    </tr>
    <tr > 
      <td align="right" valign="top" class="TableBody1">文章内容：</td>
      <td class="forumRowHighlight" style="PADDING-LEFT: 10px"><textarea id="newscontent" name="content" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"><%=htmlspecialchars(rs("content"))%></textarea>
        &nbsp;</td>
    </tr>
    <tr > 
      <td height="30" class="TableBody"></td>
      <td class="TableBody2" bgcolor="<%=m_top%>"> <input type="submit" value=" 修 改 " name="Submit" class="unnamed5" OnClick="document.form1.Content.value = frames.message.document.body.innerHTML;">
        &nbsp; <input type="reset" value=" 清 除 " name="Reset" class="unnamed5">
        &nbsp; <input type="button" value=" 返 回 " onClick="javascript:history.go(-1)" class="unnamed5"> 
        <input type="hidden" name="Content" value=""> <input type="hidden" name="linkaddress" value="<%=request.servervariables("http_referer")%>"> 
        <input type="hidden" name="newsid" value="<%=NewsID%>"> </td>
    </tr>
  </table>
</form>
<!--#include file="foot.asp"-->
</body>
</html>
