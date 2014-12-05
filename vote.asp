<!--#include file="conn.asp"-->
<%dim options,total,sql,i,answer
if request.QueryString("stype")="" then
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
		response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，谢谢！');"
		response.write"javascript:window.close();</SCRIPT>"
	else
		options=request.form("options")
		response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 
		conn.execute("UPDATE vote set answer"&options&"=answer"&options&"+1 where IsChecked=1")
	end if
end if
%><head><title>投票结果</title>
<LINK href=images/css.css rel=stylesheet>
</head>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td height="29"> 
        <table width=170 height="20" align="center" cellpadding=0 cellspacing=0>
          <tr> 
            <td width=5>&nbsp;</td>
            <td width=28> 
              <div align="center"></div>
            </td>
            <td class=hg12 valign=bottom width="123">&nbsp;<font color="#000000"><b>投票结果</b></font></td>
            <td width=12>&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="29"> 
      <table border="0" cellpadding="0" cellspacing="0" width="95%" height="48" style="border-collapse: collapse" align="center">
        <%
total=0
set rs=server.createobject("adodb.recordset")
sql="select * from vote where IsChecked=1"
rs.open sql,conn,3,3
%>
        <tr> 
          <td height="48" valign="top" colspan="3" align=left><font color="#000000"><br>
            =======================================================<br>
          </font> <font color="#000073"> <%=rs("Title")%></font> </td>
        </tr>
        <tr>
          <td valign="top">序号</td>
          <td valign="top">百比分</td>
          <td valign="top">人数</td>
        </tr>
        <%
for i=1 to 8
	if rs("Select"&i)<>"" then
		total=total+rs("answer"&i)
	end if
next
%>
        <%for i=1 to 8
	if rs("Select"&i)<>"" then
		if total=0 then
			answer=0
		else
			answer=(rs("answer"&i)/total)*100
		end if
%>
        <tr> 
          <td valign="top"><%=i%>.<%=rs("select"&i)%>：</td>
          <td valign="top"><img src=images/RSCount.gif width=<%=int(answer*2)%> height=8> 
            <%=round(answer,3)%>%</td>
          <td valign="top"><%=rs("answer"&i)%>人</td>
          <%
	end if
next
%>
        <tr>
          <td colspan="3"> 共有【<%=total%>】人参加投票<br>
            =======================================================</td>
        </tr>
      </table>
      <div align="center"></div>
      <p align="center">【<a href="javascript:window.close()">关闭窗口</a>】 
        <% rs.close     
set rs=nothing     
conn.close     
set conn=nothing %>
    </td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="23"> 
      <div align="center"></div>
      </td>
  </tr>
</table>
