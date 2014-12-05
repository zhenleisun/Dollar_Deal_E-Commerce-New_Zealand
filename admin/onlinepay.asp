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
<%dim action
action=request.QueryString("action")
'//保存信息
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.Open "select payid,paypin,payurl,npayid,npaypin,npayurl,westid,westurl,alipayid,alipaymd5,paytype,alipaytype,ipayid,ipaypin,ipayurl,ypayid,ypaypin,ypayurl,partner,py1,ems1,other,maijia from webinfo ",conn,1,3
rs("paypin")=trim(request("paypin"))
rs("payid")=request("payid")
rs("payurl")=request("payurl")
rs("npaypin")=trim(request("npaypin"))
rs("npayid")=request("npayid")
rs("npayurl")=request("npayurl")
rs("ipaypin")=trim(request("ipaypin"))
rs("ipayid")=trim(request("ipayid"))
rs("ipayurl")=trim(request("ipayurl"))
rs("ypayid")=trim(request("ypayid"))
rs("ypaypin")=trim(request("ypaypin"))
rs("ypayurl")=trim(request("ypayurl"))
rs("westid")=request("westid")
rs("westurl")=request("westurl")
rs("paytype")=request("payname")
rs("alipaytype")=request("alipaytype")
rs("alipayid")=trim(request("alipayid"))
rs("alipaymd5")=trim(request("alipaymd5"))
rs("partner")=trim(request("partner"))
rs("py1")=trim(request("py"))
rs("ems1")=trim(request("ems"))
rs("other")=trim(request("other"))
rs("maijia")=trim(request("maijia"))
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

<table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" >
  <tr> 
    <td height="25" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">在线支付设置</font></b> </td>
  </tr>
  <tr> 
    <td height="107" valign="top" >
      <table border="0" cellpadding="0" cellspacing="1"  width="100%">
        <tr>
            <td width="100%">&nbsp;</td>
        </tr>
            <tr>
            
          <td width="100%"><table width="500" border="0" align="center" cellpadding="3" cellspacing="0">
            <form name="form1" method="post" action="onlinepay.asp?action=save">
              <%set rs=server.CreateObject("adodb.recordset")
            rs.Open "select * from webinfo",conn,1,1%>
              <tr align="center" >
                <td height="22" colspan="2">网银在线支付配置 </td>
              </tr>
              <tr >
                <td height="22"><div align="left">商户ID：</div></td>
                <td width="278" height="22"><input name="payid" type="text" id="payid" size="36" value="<%=trim(rs("payid"))%>">
                </td>
              </tr>
              <tr >
                <td height="22"><div align="left">MD5密钥：</div></td>
                <td height="22"><input name="paypin" type="text" id="paypin" size="36" value="<%=trim(rs("paypin"))%>">
                </td>
              </tr>
              <tr >
                <td height="0"><div align="left">返回地址：</div></td>
                <td height="2"><input name="payurl" type="text" id="payurl" size="36" value="<%=trim(rs("payurl"))%>">
                </td>
              </tr>
              <tr >
                <td height="0">&nbsp;</td>
                <td height="3">&nbsp;</td>
              </tr>
              <tr  >
                <td colspan="2" align="center">&nbsp;</td>
              </tr>
              <tr  >
                <td colspan="2" align="center">支付宝设置</td>
              </tr>
              <tr >
                <td height="0">支付宝帐号：</td>
                <td height="22"><input name="alipayid" type="text" id="alipayid" value="<%=trim(rs("alipayid"))%>"  size="36"></td>
              </tr>
              <tr >
                <td height="2">安全校验码：</td>
                <td height="2"><input name="alipaymd5" type="text" id="alipaymd5" value="<%=trim(rs("alipaymd5"))%>" size="36"></td>
              </tr>
              <tr >
                <td height="1">合作者身份ID</td>
                <td height="1"><input name="partner" type="text" id="partner" value="<%=trim(rs("partner"))%>" size="36"></td>
              </tr>
              <tr >
                <td height="6">支付宝状态：</td>
                <td height="6">打开状态
                  <input type="radio" name="alipaytype" <% if rs("alipaytype")=0 then %> checked <%end if %> value="0">
                  关闭状态
                  <input type="radio" name="alipaytype" <% if rs("alipaytype")=1 then %> checked <%end if %> value="1">
                  //关闭状态下不显示支付宝按钮</td>
              </tr>
              <tr  >
                <td colspan="2" align="center">&nbsp;</td>
              </tr>
              <tr>
                <td height="22" colspan="2"><div align="center">
                    <input class="button" type="submit" name="Submit" value="提 交">
                  &nbsp;&nbsp;
                  <input class="button" type="reset" name="Submit2" value="恢 复">
                </div></td>
              </tr>
            </form>
          </table>
          <%rs.Close
            set rs=nothing%></td>
        </tr>
</table> 
      
    </td>
  </tr>
</table>
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
