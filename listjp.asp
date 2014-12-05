<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--商品信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="include/head.asp"-->
<TABLE cellSpacing=0 cellPadding=0 width=772 align=center border=0>
  <TBODY>
    <TR>
      <td width="1" background="image/bgbg.gif"></td>
      <TD class=b vAlign=top align=left width=764><table width="760" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
        <tr>
          <td width="100%" valign="top" bgcolor="#FFFFFF" bordercolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="4" align="center" class="table-shangxia">
              <tr>
                <td background="images/class_bg.jpg" height=50>　<img src="images/ring02.gif" width="23" height="15" align="absmiddle"> <a href=index.asp><%=webname%></a> >> 奖品信息</td>
              </tr>
            </table>
              <%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""非法访问!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""非法访问!"");location.href=""index.asp"";</script>"
end if%>
              <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from jiangpin where bookid="&request("id"),conn,1,3
if rs.recordcount=0 then 
%>
              <table width="370" border="0" cellspacing="0" cellpadding="5" align="center">
                <tr>
                  <td align=center>调用奖品错误</td>
                </tr>
              </table>
              <%else%>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td valign="top" width="40%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><table align=center cellspacing=0 cellpadding=0 width=220 height=220 border=0>
                            <tbody>
                              <tr>
                                <td background=images/cla.gif align=center><%if rs("bookpic")<>"" then 
response.write "<a href="&trim(rs("bookpic2"))&" ><img src="&trim(rs("bookpic"))&" width=150 border=0></a>"
else
response.write "<img src=images/emptybook.gif width=150 border=0>"
end if%>
                                </td>
                              </tr>
                            </tbody>
                        </table></td>
                      </tr>
                  </table></td>
                  <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td>市 场 价：<%=rs("shichangjia")%> 元，所需积分：<%=rs("point")%></td>
                      </tr>
                      <tr>
                        <td>奖品名称：<%=rs("bookname")%></td>
                      </tr>
                      <tr>
                        <td>奖品说明：<%=rs("bookcontent")%></td>
                      </tr>
                  </table></td>
                </tr>
                <%
		rs.close
		set rs=nothing
		end if%>
                <%
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from review where bookid="&request("id")&" order by pinglundate desc",conn,1,1
		j=rs.recordcount
		'if i>5 then j=5
		for i=1 to j
		%>
                <%	
	rs.movenext
	next
	%>
            </table></td>
        </tr>
      </table></TD>
        
      <td width="1" background="image/bgbg.gif"></td>
    </TR>
  </TBODY>
</TABLE>
<!--#include file="include/foot.asp"-->
</body>
</html>