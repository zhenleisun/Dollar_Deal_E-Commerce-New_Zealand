<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<%dim action,songid
songid=request.QueryString("id")
if songid<>"" then
if not isnumeric(songid) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
action=request.QueryString("action")
set rs=server.CreateObject("adodb.recordset")
select case action
case "songhuosave"
rs.open "select * from deliver where songid="&songid,conn,1,3
rs("subject")=trim(request("subject2"))
rs("songidorder")=request("songidorder2")
rs("jsmoney")=request("jsmoney")
rs("fangshi")=0
rs.update
rs.close
response.write "<script language=javascript>alert('送货方式成功修改！');history.go(-1);</script>"
response.End
case "songhuoadd"
rs.open "select * from deliver",conn,1,3
rs.addnew
rs("subject")=trim(request("subject"))
rs("songidorder")=request("songidorder")
rs("jsmoney")=request("jsmoney")
rs("fangshi")=0
rs.update
rs.close
response.write "<script language=javascript>alert('新的送货方式成功添加！');history.go(-1);</script>"
response.End
case "songhuodel"
conn.execute "delete from deliver where songid="&songid
response.redirect "manageremit.asp?action=songhuo"
case "zhifudel"
conn.execute "delete from deliver where songid="&songid
response.redirect "manageremit.asp?action=zhifu"
case "zhifusave"
rs.open "select * from deliver where songid="&songid,conn,1,3
rs("subject")=trim(request("subject"))
rs("songidorder")=request("songidorder")
rs("fangshi")=1
rs.update
rs.close
response.write "<script language=javascript>alert('支付方式成功修改！');history.go(-1);</script>"
response.End
case "zhifuadd"
rs.open "select * from deliver",conn,1,3
rs.addnew
rs("subject")=trim(request("subject"))
rs("songidorder")=request("songidorder")
rs("fangshi")=1
rs.update
rs.close
response.write "<script language=javascript>alert('新的支付方式成功添加！');history.go(-1);</script>"
response.End
end select
set rs=nothing
%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr>
<td  align="center" background="../images/admin_bg_1.gif"><b><a href="manageremit.asp?action=songhuo"><font color="#ffffff">点击>修改送货方式</font></a></b></td>
<td  align="center" background="../images/admin_bg_1.gif"><b><a href="manageremit.asp?action=zhifu"><font color="#ffffff">点击>修改支付方式</font></a></b></td>
</tr>
<tr> 
<td  colspan="2">
<%select case action
	case "songhuo"%>
                              
      <table class="tableBorder"  width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr> 
    <td class="forumRowHighlight"><div align="center">送货方式</div></td>
    <td class="forumRowHighlight"><div align="center"><span lang="zh-cn">加收金额</span></div></td>
    <td class="forumRowHighlight"><div align="center">排 序</div></td>
    <td class="forumRowHighlight"><div align="center">操作</div></td>
  </tr>
  <%dim i,j
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from deliver where fangshi=0 order by songidorder",conn,1,1
		i=rs.recordcount
		do while not rs.eof%>
  <tr> 
    <form name="form1" method="post" action="manageremit.asp?action=songhuosave&id=<%=rs("songid")%>">
      <td class="forumRowHighlight"> <div align="center"> 
          <input name="subject2" type="text" id="subject2" size="14" value=<%=trim(rs("subject"))%>>
        </div></td>
      <td class="forumRowHighlight" width="127"> <p align="center"> 
          <input name="jsmoney" type="text" id="jsmoney" size="4" value=<%=rs("jsmoney")%> onkeypress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
          　</td>
      <td class="forumRowHighlight"> <div align="center"> 
          <input name="songidorder2" type="text" id="songidorder2" size="2" value=<%=rs("songidorder")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
        </div></td>
      <td class="forumRowHighlight"> <div align="center"> 
          <input class="button" type="submit" name="Submit" value="修 改">
          &nbsp;<a href="manageremit.asp?action=songhuodel&id=<%=rs("songid")%>" onClick="return confirm('您确定进行删除操作吗？')">删除</a> 
        </div></td>
    </form>
  </tr>
  <%rs.movenext
		loop
		rs.close
		set rs=nothing%>
  <tr> 
    <form name="form2" method="post" action="manageremit.asp?action=songhuoadd">
      <td class="forumRowHighlight"> <div align="center"> 
                <input name="subject" type="text" id="subject" size="14">
        </div></td>
      <td class="forumRowHighlight" width="127"> <p align="center"> 
          <input name="jsmoney" type="text" id="jsmoney" size="4" onkeypress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
          　</td>
      <td class="forumRowHighlight"> <div align="center"> 
                <input name="songidorder" type="text" id="songidorder" value=<%=i+1%> size="2" onKeyPress= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
        </div></td>
      <td class="forumRowHighlight"> <div align="center"> 
          <input class="button" type="submit" name="Submit3" value="添 加">
        </div></td>
    </form>
  </tr>
</table>
                              <%case "zhifu"%>
                              <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" >
                                <tr >
                                  <td width="40%" align="center" bgcolor="fbc2c2" >支付方式</td>
                                  <td width="20%" align="center" bgcolor="fbc2c2" >排 序</td>
                                  <td width="40%" align="center" bgcolor="fbc2c2" >操 作</td>
                                </tr>
                                <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from deliver where fangshi=1 order by songidorder",conn,1,1
		  j=rs.recordcount
		  do while not rs.eof%>
                                <tr> 
                                  <form name="form1" method="post" action="manageremit.asp?action=zhifusave&id=<%=rs("songid")%>">
                                    <td  align="center">
									<input name="subject" type="text" id="subject" size="14" value=<%=trim(rs("subject"))%>>									</td>
                                    <td  align="center">
									<input name="songidorder" type="text" id="songidorder" size="2" value=<%=rs("songidorder")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">                                    </td>
                                    <td align="center"  STYLE='PADDING-LEFT: 20px'><input type="submit" name="Submit2" value="确 认">
									&nbsp;<a href="manageremit.asp?action=zhifudel&id=<%=rs("songid")%>" onClick="return confirm('您确定进行删除操作吗？')"><font color="#FF0000">删除</font></a>                                    </td>
                                  </form>
                                  <%rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                                </tr>
								<tr>
								<td colspan="3"  align="center" bgcolor="fbc2c2" >添加支付方式</td>
								</tr>
                                <tr> 
                                  <form name="form1" method="post" action="manageremit.asp?action=zhifuadd">
                                    <td  align="center">
									<input name="subject" type="text" id="subject" size="14">                                    </td>
                                    <td  align="center">
									<input name="songidorder" type="text" id="songidorder" value=<%=j+1%> size="2" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
<td  STYLE='PADDING-LEFT: 20px'>
<input type="submit" name="Submit32" value="添 加"></td>
</form>
</tr>
                                <tr>
                                  <td height="30" colspan="3"  align="center"><font color="ff0000">支付方式中的“货到付款”“预存款支付”“在线支付”三种方式不能改动，不能删除。</font></td>
                                </tr>
</table>
<%end select%>
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
