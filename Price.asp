<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>

<title><%=webname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="<%=des%>">
<meta name="keywords" content="<%=keya%>">

<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<!--#include file="include/head.asp"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="39" valign="top"><table width="100%" height="184" border="0" cellpadding="0" cellspacing="0">
      <tr>
          <td height="30" align="right" valign="top">&nbsp;</td>
      </tr>
      <tr>
        <td height="90"><div align="right"></div></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table></td>
	<td width="190" valign="top" style="BORDER-bottom:<%=skincolor%> 1px solid;BORDER-left:<%=skincolor%> 1px solid; BORDER-right:<%=skincolor%> 1px solid;">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="userinfo.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td><!--#include file="searc.asp"--><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td></td>
      </tr>
    </table></td>
    <td width="700" valign="top"><table cellspacing=0 cellpadding=0 width=100% align=center border=0>
      <tbody>
        <tr>
          <td class=b valign=top align=left width=100% >
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
                <tr>
                  <td width="100%" align="center" valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="left" height="28" background="images/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                          <a href=index.asp><%=webname%></a> >> <a href="price.asp">报价中心</a> 
                        </td>
                      </tr>
                      <tr bgcolor="#ffffff">
                        <td  valign="top"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
							
							<%
							anid=trim(request("anid"))
							%>
                              <td><br>请选择需要查看报价的类别： 
                                <table width="100%" border="0" cellpadding="0" cellspacing="0" dwcopytype="CopyTableCell">
          <tr>
                <td align=center><span class="style2">
                 <a href="price.asp">全部商品</a> <% set rs=server.CreateObject("adodb.recordset")
rs.open "select * from bsort order by anclassidorder asc",conn,1,1
do while not rs.eof
		  response.write "||  <A href=?anid="&rs("anclassid")&">"&trim(rs("anclass"))&"</a>  "
		 rs.movenext
loop
rs.close
set rs=Nothing
		 %>
                  </span></td>
          </tr>
        </table><%
Const MaxPerPage=20
dim totalPut   
dim CurrentPage
dim TotalPages
dim j
dim sql
if Not isempty(request("page")) then
currentPage=Cint(request("page"))
else
currentPage=1
end if 
set rs=server.CreateObject("adodb.recordset")
if anid<>"" then
rs.open "select * from products where  anclassid="&anid&" and zhuangtai=0 order by adddate desc",conn,1,1
else
select case selectm
case ""
rs.open "select * from products  where zhuangtai=0 order by adddate desc",conn,1,1
case "0"
rs.open "select * from products where zhuangtai=0  order by adddate desc",conn,1,1
case "shopid"


end select
end if
if err.number<>0 then
response.write "暂无相关数据！"
end if
if rs.eof And rs.bof then
Response.Write "<p align='center'>暂无相关数据！</p>"
else
totalPut=rs.recordcount
if currentpage<1 then
currentpage=1
end if
if (currentpage-1)*MaxPerPage>totalput then
if (totalPut mod MaxPerPage)=0 then
currentpage= totalPut \ MaxPerPage
else
currentpage= totalPut \ MaxPerPage + 1
end if
end if
if currentPage=1 then
showContent
showpage totalput,MaxPerPage,"Price.asp"
else
if (currentPage-1)*MaxPerPage<totalPut then
rs.move  (currentPage-1)*MaxPerPage
dim shopmark
shopmark=rs.bookmark
showContent
showpage totalput,MaxPerPage,"Price.asp"
else
currentPage=1
showContent
showpage totalput,MaxPerPage,"Price.asp"
end if
end if
end if
sub showContent
dim i
i=0
%>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            </td>
          <td ></td>
        </tr>
      </tbody>
    </table>
      <form name="form1" method="post" action="">
        <table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#999999">
          <tr bgcolor="#C7D3E6"> 
            <td height="28" align="center" bgcolor="#F7F7F7">商品序号</td>
            <td height="28" align="center" bgcolor="#F7F7F7">商品名称</td>
            <td height="28" align="center" bgcolor="#F7F7F7">市场价</td>
            <td height="28" align="center" bgcolor="#F7F7F7">会员价</td>
            <td height="28" align="center" bgcolor="#F7F7F7">当前库存</td>
            <td height="28" align="center" bgcolor="#F7F7F7">商品品牌</td>
            <td height="28" align="center" bgcolor="#F7F7F7">规格</td>
            <td height="28" align="center" bgcolor="#F7F7F7">规格参数</td>
          </tr>
          <%do while not rs.eof%>
          <tr bgcolor="#F0F3F8"> 
            <td height="18" align="center" bgcolor="#FFFFFF"><%=rs("bookid")%></td>
            <td height="18" align="left" bgcolor="#FFFFFF" style='PADDING-LEFT: 10px'><img src="images/price_ico.png" width="11" height="11" /> 
              <a href=products.asp?id=<%=rs("bookid")%> target="_blank"> 
              <%
	if len(trim(rs("bookname")))>28 Then
	response.write left(trim(rs("bookname")),28)&".."
	Else
	response.write trim(rs("bookname"))
	End If
	%>
              </a></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><s><%=rs("shichangjia")%></s></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><%=rs("huiyuanjia")%></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><%=rs("kucun")%></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><%=rs("pingpai")%></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><%=rs("isbn")%></td>
            <td height="18" align="center" bgcolor="#FFFFFF"><a href=products.asp?id=<%=rs("bookid")%> target="_blank"> 查看详细</a></td>
          </tr>
          <%
  i=i+1
  if i>=MaxPerPage then Exit Do
  rs.movenext
  Loop
  rs.close
  set rs=Nothing
  %>
        </table>
      </form>
	  
	  
	  <%
End Sub
Function showpage(totalnumber,maxperpage,filename)  
Dim n

If totalnumber Mod maxperpage=0 Then  
n= totalnumber \ maxperpage  
Else
n= totalnumber \ maxperpage+1  
End If

Response.Write "<form method=Post action="&filename&"?selectm="&selectm&"&selectkey="&selectkey&" >"  
Response.Write "<p align='center'> "  
If CurrentPage<2 Then  
Response.Write "首页 上一页 "  
Else  
Response.Write "<a href="&filename&"?page=1&anid="&anid&"&selectm="&selectm&"&selectkey="&selectkey&">首页</a> "  
Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&anid="&anid&"&selectm="&selectm&"&selectkey="&selectkey&">上一页</a> "  
End If

If n-currentpage<1 Then  
Response.Write "下一页 尾页"  
Else  
Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&anid="&anid&"&selectm="&selectm&"&selectkey="&selectkey&">"  
Response.Write "下一页</a> <a href="&filename&"?page="&n&"&anid="&anid&"&selectm="&selectm&"&selectkey="&selectkey&">尾页</a>"  
End If  
Response.Write " 页次："&CurrentPage&"/"&n&"页 "  
Response.Write " 共有"&totalnumber&"种商品 " 
Response.Write "转到：<input type='text' name='page' size=2 maxlength=10 value="&currentpage&">"  
Response.Write "&nbsp;<input type='submit' value='GO' name='cndok'></form>"  
End Function  
%> </td>
	<td width="68"style="BORDER-left:<%=skincolor%> 1px solid;">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
