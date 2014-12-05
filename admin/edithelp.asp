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
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
select case action
case ""%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="0" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">网站帮助信息设置</font></b></td>
</tr>
<tr><td >
                              <table width="90%" border="0" align="center" cellpadding="2" cellspacing="1">
                                <tr  align="center"> 
                                  <td width="27%"> <a href="edithelp.asp?action=huikuanfangshi">付款方式</a></td>
                                  <td width="34%"> <a href="edithelp.asp?action=gouwuliucheng">购物流程</a></td>
                                  <td width="39%"><a href="edithelp.asp?action=lxwm">联系我们</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=jiaoyitiaokuan">交易条款</a></td>
                                  <td> <a href="edithelp.asp?action=regtiaoyue">注册条约</a></td>
                                  <td> <a href="edithelp.asp?action=shiyongfalv">版权声明</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=yunshushuoming">配送须知</a></td>
                                  <td> <a href="edithelp.asp?action=baomi">保密和安全</a></td>
                                  <td> <a href="edithelp.asp?action=shouhoufuwu">商品销售/售后</a></td>
                                </tr>
                                <tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=songhuofeiyong">送货方式和费用</a></td>
                                  <td> <a href="edithelp.asp?action=gongzuoshijian">正常上班时间</a></td>
                                  <td> <a href="edithelp.asp?action=about">关于我们</a></td>
                                </tr>
								<tr  align="center"> 
                                  <td> <a href="edithelp.asp?action=baozhuangbiaozhun">个人包装标准</a></td>
                                  <td> <a href="edithelp.asp?action=baozhuangshuoming">包装说明</a></td>
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
if action="huikuanfangshi" then response.write "付 款 方 式"
if action="gouwuliucheng" then response.write "购 物 流 程"
if action="regtiaoyue" then response.write "更改注册条约"
if action="yunshushuoming" then response.write "配 送 须 知"
if action="baomi" then response.write "保 密 和 安 全"
if action="shouhoufuwu" then response.write "商品销售和售后服务"
if action="songhuofeiyong" then response.write "送货方式及费率"
if action="gongzuoshijian" then response.write "我们的工作时间"
if action="about" then response.write "关 于 本 站"
if action="lxwm" then response.write "联 系 我 们"
if action="jiaoyitiaokuan" then response.write "交 易 条 款"
if action="baozhuangbiaozhun" then response.write "个人包装标准"
if action="baozhuangshuoming" then response.write "包 装 说 明"
if action="shiyongfalv" then response.write "适用法律和版权声明"
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
											<input type="button" value=" 返 回 " onClick="javascript:history.go(-1)" class="unnamed5" name="button">&nbsp;
											<input type="submit" value=" 修 改 " name="Submit" class="unnamed5" onClick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp;
											<input type="reset" value=" 清 除 " name="Reset" class="unnamed5">
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
