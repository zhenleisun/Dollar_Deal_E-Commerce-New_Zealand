<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script LANGUAGE='javascript'>alert('���糬ʱ��������û�е�¼���¼');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
 if request.QueryString("action")="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from news",conn,1,3
rs.addnew
rs("newsname")=trim(request("newsname"))
rs("addname")=trim(request("addname"))

Content=Request.Form("Content")
rs("newscontent")=content
rs("adddate")=now()
rs("viewcount")=0
rs.update
rs.close
set rs=nothing
session("content")=""
response.write "<script language=javascript>alert('��ӳɹ���');window.location.href='editnews.asp';</script>"
response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
 
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

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td  align="center" background="../images/admin_bg_1.gif" height="25"><b><font color="#ffffff">�����վ����</font></b></td>
</tr>
<tr> 
<td valign="top"> 
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FBF4F4">
        <form name="form1" method="post" action="addnews.asp?action=save" onSubmit="return checkdata()" onReset="return ResetForm();">
          <tr > 
            <td width="30%" align="right">�������⣺</td>
            <td width="70%" style="PADDING-LEFT: 10px"> <input name="newsname" type="text" id="newsname" size="30"> 
            </td>
          </tr>
          <tr > 
            <td align="right">�� �� �ˣ�</td>
            <td style="PADDING-LEFT: 10px"> <input name="addname" type="text" id="addname" size="30"> 
            </td>
          </tr>
          <tr > 
            <td align="right" valign="top">�������ݣ�</td>
            <td bgcolor="#FFFFFF" class="forumRowHighlight"style="PADDING-LEFT: 10px"><textarea id="newscontent" name="Content" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"></textarea>
               
              &nbsp;</td>
          </tr>
          <tr > 
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px"> <input type="submit" name="Submit" value="�ύ����" onClick="document.form1.Content.value = frames.message.document.body.innerHTML;"> 
              <input onClick="ClearReset()" type=reset name="Clear" value="������д"> 
              <input type="hidden" name="Content" value=""> </td>
          </tr>
        </form>
      </table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
