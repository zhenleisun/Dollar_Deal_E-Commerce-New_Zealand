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
%><%


Dim htmlData

htmlData = Request.Form("Content")

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
			id : 'Content',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>

<html>
<head>
<title>VIP˵��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link href="../images/css.css" rel="stylesheet" type="text/css">
<body>
<%
set rs=server.CreateObject("adodb.recordset")
rs.Open "select vipsq from webinfo",conn,1,3
%>
<form name="form1" method="post" action="?action=save" OnSubmit="return checkdata()" onReset="return ResetForm();">
  <table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
    <tr> 
      <th  background="../images/admin_bg_1.gif" class="tableHeaderText">VIP����˵��ҳ��</th>
    </tr>
    <tr> 
      <td align=center class="forumRowHighlight"style="PADDING-LEFT: 10px"><textarea id="Content" name="Content" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"><%=htmlspecialchars(rs("vipsq"))%></textarea> 
        &nbsp;</td>
    </tr>
    <tr> 
      <td align="center" class="forumRowHighlight"> <input type="submit" value=" �� �� " name="Submit" class="button" onClick="document.form1.Content.value = frames.message.document.body.innerHTML;"> 
      </td>
    </tr>
    <tr> 
      <td align=center class="forumRowHighlight"style="PADDING-LEFT: 10px">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
<%
if request("action")="save" then

html=request("html")
action=request.QueryString("action")
Content=Request.Form("Content")


set rs=server.CreateObject("adodb.recordset")
rs.Open "select vipsq from webinfo",conn,1,3
rs("vipsq")=content
rs.Update

response.Write "<script language=javascript>alert('VIP����ҳ���޸ĳɹ�');window.location.href='VipExplain.asp';</script>"
response.End
end if 
%>