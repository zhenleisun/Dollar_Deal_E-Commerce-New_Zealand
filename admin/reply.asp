<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End

end if
%>
<%
set rs=server.createobject("adodb.recordset")
rs.open "select * from guestbook where id="&request("id"),conn,3,2
if request("action")<>"save" then
%>
<html>
<head>
<title>回复</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<form action="reply.asp?action=save&id=<%=request("id")%>" method="post" name="form" id="form">
  <table width="88%" border="1" align="center" cellpadding="0" cellspacing="5" bordercolor="#0033FF" class=log_table>
    <tr> 
      <td align="center"> <table  width="100%" border="0" class=log_titlewidth="100%">
          <tr> 
            <td>&nbsp;&nbsp;&nbsp;留言内容：</td>
          </tr>
          <tr> 
            <td> 
              <%
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
	IF Err.Number <>0 Then
		unHtml= "HTML转换中出错请联系管理员<br>"
		Err.Clear
	End IF
End Function
%>
              &nbsp;&nbsp;&nbsp;<%=unHtml(rs("content"))%></td>
          </tr>
          <tr> 
            <td><strong>&nbsp;&nbsp;&nbsp;管 理 员 回 复 </strong></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="65%">&nbsp;&nbsp;&nbsp;回复<font color="#990033"><%=rs("name")%></font>的留言</td>
            <td rowspan="2"><div align="right"></div></td>
          </tr>
          <tr> 
            <td>&nbsp;&nbsp;&nbsp;不超过255个字符！</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td valign="middle">&nbsp;&nbsp;&nbsp;回复内容： <br> &nbsp;&nbsp;&nbsp; <textarea name="reply" cols="58" rows="8" class="wenbenkuang" id="reply"><%=rs("reply")%></textarea> 
      </td>
    </tr>
    <tr>
      <td valign="middle"><input type="radio" name="Online" value="0" <%if rs("Online")="0" then%>checked<%end if%>>
        隐藏
        <input type="radio" name="Online" value="1" <%if rs("Online")="1" then%>checked<%end if%>>
        公开 </td>
    </tr>
    <tr> 
      <td height="30"> &nbsp;&nbsp;&nbsp; <input type="submit" class="go-wenbenkuang" value="回复" name="button"> 
        &nbsp;&nbsp;
        <input name="Submit2" type="reset" class="go-wenbenkuang" id="Submit2" value="重置"> 
      </td>
    </tr>
  </table>
  </form>
<%
else

rs("reply")=Server.HTMLEncode(request.form("reply"))
rs("Online")=request("Online")
rs.update
rs.close
set rs=nothing
response.redirect "viewreturn.asp"

end if
%>
</body>
</html>