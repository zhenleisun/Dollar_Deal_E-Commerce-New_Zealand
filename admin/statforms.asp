
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
'day1=trim(Request.Form("year1"))&"-"&trim(Request.Form("month1"))&"-"&trim(Request.Form("day1"))
'day2=trim(Request.Form("year2"))&"-"&trim(Request.Form("month2"))&"-"&trim(Request.Form("day2"))
day1="#"&trim(Request.Form("month1"))&"-"&trim(Request.Form("day1"))&"-"&trim(Request.Form("year1"))&"#"
day2="#"&trim(Request.Form("month2"))&"-"&trim(Request.Form("day2"))&"-"&trim(Request.Form("year2"))&"#"
%>
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" height="24" bgcolor="#cccccc"><b><font color="#ffffff">统计报表查看</font></b></td>
</tr>
<tr>
<%
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from products order by bookid",conn,1,1
	%>
<td height="107"> 
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center" > 
<td width="11%">商品ID</td>
<td width="41%">商品名称</td>
<td width="17%">商品单位</td>
<td width="14%">销售数量</td>
<td width="17%">销售额</td>
</tr>
                                <%
									z_xs_sl=0
									z_xs_je=0
									while not rs.eof
									set rs2=server.createobject("adodb.recordset")
									rs2.open "select * from orders where (datediff('d',fhsj,"&day2&")>=0)  and (datediff('d',"&day1&",fhsj)>=0)  and (bookid="&rs("bookid")&")",conn,1,1
									'rs2.open "select * from orders where (datediff('d',fhsj,"&day2&")>=0)  and (bookid="&rs("bookid")&")",conn,1,1
									'response.write "select * from orders where  datediff('d',fhsj,"&day2&")>=0 and  datediff('d',"&day1&",fhsj)<=0 and bookid="&rs("bookid")
									'response.end
									xs_sl=0
									xs_je=0
									while not rs2.eof
										xs_sl=xs_sl+rs2("bookcount")
										xs_je=xs_je+rs2("zonger")
										rs2.movenext
									wend
									rs2.close
									set rs2=nothing
								%>
<tr > 
<td align="center" width="11%"><%=rs("bookid")%></td>
<td width="41%" STYLE='PADDING-LEFT: 10px'><%=rs("bookname")%></td>
<td align="center" width="17%"><%=rs("bookchuban")%></td>
<td align="center" width="14%" ><%=xs_sl%></td>
<td align="center" width="17%"><%=xs_je%></td>
</tr>
                                <%
									z_xs_sl=z_xs_sl+xs_sl
									z_xs_je=z_xs_je+xs_je
								rs.movenext
								wend
								rs.close
								set rs=nothing
								%>
<tr > 
<td colspan="5" align="center"> 
<input type="button" name="Submit" value="返回" onClick="javascript:history.go(-1)">
</td>
</tr>
</table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
