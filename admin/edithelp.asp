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
action=request("action")
set rs=server.createobject("adodb.recordset")
sql="select "&Action&" from webinfo "
rs.open sql,conn,1,1
Ccontent=rs(0)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
 
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
 <%


Dim htmlData

htmlData = Request.Form("contentt")

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
			id : 'contentt',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>
<%dim action
action=request.QueryString("action")
if InStr(Action,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
select case action
case ""%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="0" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">��վ������Ϣ����</font></b></td>
</tr>
<tr><td >
                              <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1">
                                <tr  align="center"> 
                                  <td width="27%"> <a href="edithelp.asp?action=huikuanfangshi">���ʽ</a></td>
                                  <td width="34%"> <a href="edithelp.asp?action=gouwuliucheng">��������</a></td>
                                  <td width="39%"><a href="edithelp.asp?action=lxwm">��ϵ����</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=jiaoyitiaokuan">��������</a></td>
                                  <td> <a href="edithelp.asp?action=regtiaoyue">ע����Լ</a></td>
                                  <td> <a href="edithelp.asp?action=shiyongfalv">��Ȩ����</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=yunshushuoming">������֪</a></td>
                                  <td> <a href="edithelp.asp?action=baomi">���ܺͰ�ȫ</a></td>
                                  <td> <a href="edithelp.asp?action=shouhoufuwu">��Ʒ����/�ۺ�</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=songhuofeiyong">�ͻ���ʽ�ͷ���</a></td>
                                  <td> <a href="edithelp.asp?action=gongzuoshijian">�����ϰ�ʱ��</a></td>
                                  <td> <a href="edithelp.asp?action=about">��������</a></td>
                                </tr>
								<tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=baozhuangbiaozhun">���˰�װ��׼</a></td>
                                  <td> <a href="edithelp.asp?action=baozhuangshuoming">��װ˵��</a></td>
                                  <td> </td>
                                </tr>
                                </table>
                            </td>
                          </tr>
                        </table>
<%case else%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="0" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">
	<%
if action="huikuanfangshi" then response.write "�� �� �� ʽ"
if action="gouwuliucheng" then response.write "�� �� �� ��"
if action="regtiaoyue" then response.write "����ע����Լ"
if action="yunshushuoming" then response.write "�� �� �� ֪"
if action="baomi" then response.write "�� �� �� �� ȫ"
if action="shouhoufuwu" then response.write "��Ʒ���ۺ��ۺ����"
if action="songhuofeiyong" then response.write "�ͻ���ʽ������"
if action="gongzuoshijian" then response.write "���ǵĹ���ʱ��"
if action="about" then response.write "�� �� �� վ"
if action="lxwm" then response.write "�� ϵ �� ��"
if action="jiaoyitiaokuan" then response.write "�� �� �� ��"
if action="baozhuangbiaozhun" then response.write "���˰�װ��׼"
if action="baozhuangshuoming" then response.write "�� װ ˵ ��"
if action="shiyongfalv" then response.write "���÷��ɺͰ�Ȩ����"
%>
</font></b></td>
</tr><tr><td>
                              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                                <form name="form1" method="post" action="savehelp.asp?action=<%=action%>" OnSubmit="return checkdata()" onReset="return ResetForm();">
                                  <tr> 
                                    <td  align="center">
                                        <table width="500" border="0" cellspacing="2" cellpadding="2">
                                          <tr>
										  
                  <td><br>
                  </td>
                                            </tr>
                                            <tr>
                                            <td align="left"> <textarea id="contentt" name="ccontent" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"><%=ccontent%></textarea>
                                            </td>
                                          </tr>
                                          <tr> 
                                            <td align="center">
											<input type="button" value=" �� �� " onClick="javascript:history.go(-1)" class="unnamed5" name="button">&nbsp;
											<input type="submit" value=" �� �� " name="Submit" class="unnamed5" onClick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp;
											<input type="reset" value=" �� �� " name="Reset" class="unnamed5">
											<input type="hidden" name="Content" value="">
                                            </td>
                                          </tr>
                                        </table>
                                    </td>
                                  </tr>
                                </form>
                              </table>
                            </td>
                          </tr>
                        </table>
<%end select%>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
