<html>
<head>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<title><%=webname%>--<%
	leixing=lcase(trim(request("lx")))
	leixing=replace(leixing,"'","")
	select case leixing
	case "big"
		response.write "High class"
	case "small"
		response.write "Small class"
	case "tejia"
		response.write "Specials"
	case "news"
		response.write "New product"
	case else
		response.write "Recommend product"					  
	end select%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
	  <tr>
        <td><!--#include file="searc.asp"--></td>
      </tr>
	  <tr>
        <td><!--#include file="shopcart.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td><!--#include file="include/selltop.asp" --></td>
      </tr>
	  <tr>
        <td height="5" background="images/leftbg.jpg"></td>
      </tr> 
	  <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>      
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table width="771" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
      <tr>
        <td valign="top" align="center" bgcolor="#FFFFFF"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
            <%leixing=lcase(trim(request("lx")))
			if InStr(leixing,"'")>0 then
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			response.end
			end if
			'leixing=replace(leixing,"'","")
			select case leixing
			case "big"
			anclassid=trim(request("anid"))
			if not isnumeric(anclassid) then 
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			response.end
			else
			if not isinteger(anclassid) then
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			else
			set rs=server.createobject("adodb.recordset")
			rs.open "select * FROM bsort where anclassid="&anclassid,conn,1,1
			anclassname=rs("anclass")
			%>
            <tr>
              <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td valign="top" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <%=anclassname%></td>
                          </tr>
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
							  <tr>
								<td width="1%"></td>
								<td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
							  </tr>
						  </table></td>
                        </tr>
                        
                    </table></td>
                  </tr>
				  <tr>
                    <td height="10"></td>
                  </tr>
                  <tr>
                    <td><table  width="98%"  border="0" align="center" cellpadding="1" cellspacing="1"  bgcolor="#E1E1E1">
                        <tr>
                          <td width="110" align="center" bgcolor="#FFFFFF"><%=anclassname%></td>
                          <td bgcolor="#FFFFFF"><table border="0" cellpadding="0" cellspacing="0">
                              <%
								set rs_s=server.CreateObject("adodb.recordset")
								rs_s.open "select * from ssort where anclassid="&rs("anclassid")&" order by nclassidorder",conn,1,1
								if rs_s.recordcount=0 then 
								%>
                              <tr>
                                <td height="20" align="center" colspan="5">No small class</td>
                              </tr>
                              <%
								else
								i=0
								while not rs_s.eof
								%>
                              <tr style="BORDER-bottom:#CC0033 1px solid;">
                                <td width="116"  height="24" align="left"><img src="images/body/diand.gif">&nbsp;&nbsp;<a href="class.asp?lx=small&anid=<%=rs("anclassid")%>&nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a>
                                        <%rs_s.movenext
										if rs_s.eof then 
										response.write " "
										else
										%></td>
                                <td width="116" align="left" ><img src="images/body/diand.gif">&nbsp;<a href="class.asp?lx=small&anid=<%=rs("anclassid")%>&nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a>
                                        <%rs_s.movenext
											if rs_s.eof then 
											response.write " "
											else
										%></td>
                                <td width="116" align="left" ><img src="images/body/diand.gif">&nbsp;&nbsp;<a href="class.asp?lx=small&anid=<%=rs("anclassid")%>&nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a>
                                        <%rs_s.movenext
										if rs_s.eof then 
										response.write " "
										else
										%></td>
                                <td width="116" align="left" ><img src="images/body/diand.gif">&nbsp;&nbsp;<a href="class.asp?lx=small&anid=<%=rs("anclassid")%>&nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a>
                                        <%rs_s.movenext
										if rs_s.eof then 
										response.write " "
										else
										%></td>
                                <td width="131" align="left" ><img src="images/body/diand.gif">&nbsp;&nbsp;<a href="class.asp?lx=small&anid=<%=rs("anclassid")%>&nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a> </td>
                                <%
								rs_s.movenext
								end if
								end if
								end if
								end if
								
								wend
								end if
								end if
								end if
								%>
                              </tr>
                          </table></td>
                        </tr>
                    </table></td>
                  </tr>
              </table></td>
            </tr>
			<tr>
              <td height="8"></td>
            </tr>
            <%
			case "small"
			anclassid=request("anid")
			if not isnumeric(anclassid) then 
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			response.end
			else
			if not isinteger(anclassid) then
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			else
				nclassid=request("nid")
				if not isnumeric(nclassid) then 
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			response.end
			else
			if not isinteger(nclassid) then
			response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
			else
			set rs=server.createobject("adodb.recordset")
			rs.open "select * FROM bsort where anclassid="&anclassid,conn,1,1
			anclassname=rs("anclass")
			rs.close
			rs.open "select * FROM ssort where nclassid="&nclassid,conn,1,1
			nclassname=rs("nclass")
			rs.close
			%>
			<tr>
              <td height="28" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> <a href=class.asp?lx=big&anid=<%=anclassid%>><%=anclassname%></a> >> <%=nclassname%></td>
            </tr>
			<tr>
              <td height="8"></td>
            </tr>
			<tr>
              <td height=50><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%"></td>
                    <td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
                  </tr>
              </table></td>
            </tr>
			<tr>
              <td height="8"></td>
            </tr>
            <%
			end if 
			end if
			end if
			end if
			case "tejia"%>
            <tr>
              <td height=28 valign="middle" background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Special promotions</td>
            </tr>
            <tr>
              <td height="8"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%">&nbsp;</td>
                    <td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height=10></td>
            </tr>
            <%case "hot"%>
            <tr>
              <td  height=28 background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Hot commodity</td>
            </tr>
            <tr>
              <td height="8"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%">&nbsp;</td>
                    <td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height=10></td>
            </tr>
            <%case "news"%>
            <tr>
              <td height=28 background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> New top</td>
              </tr>
            <tr>
              <td height="8"></td>
            </tr>
            <tr>
              <td height=50><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%"></td>
                    <td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height=10></td>
            </tr>
            <%case else%>
            <tr>
              <td height=28 background="images/body/pdbg01.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> New top</td>
              </tr>
            <tr>
              <td height="8"></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1%"></td>
                    <td width="99%"><a href="#" target="_blank"><img src="<%=ad8%>" border="0" width="99%" height="180" ></a></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td height=10></td>
            </tr>
            <%end select%>
          </table>
		  <%set rs=server.createobject("adodb.recordset")
					xid=request("xid")
					order=request("order")
					
					if order="PriceUp" then
					 orderby="order by huiyuanjia asc"
					 
					  elseif order="PriceDown" then
					  orderby="order by huiyuanjia desc"
					  
					 elseif order="ViewsUp" then
							
					  orderby="order by liulancount asc"
					  elseif order="ViewsDown" then
					   orderby="order by liulancount desc"
					 elseif order="Addup" then
						 orderby="order by adddate desc"
					 elseif order="SaleUp" then
						 orderby="order by chengjiaocount asc"
						 elseif order="SaleDown" then
						 orderby="order by chengjiaocount desc"
					  else
					orderby="order by adddate desc"
					 end if 
					
					if leixing="big" then
						rs.open "select * from products where anclassid="&anclassid&"and zhuangtai=0  "&orderby&"",conn,1,1
					elseif leixing="small" then
					rs.open "select * from products where anclassid="&anclassid&" and nclassid="&nclassid&" and zhuangtai=0 "&orderby&"",conn,1,1
					elseif leixing="hot" then
						rs.open "select * from products where   zhuangtai=0 "&orderby&"",conn,1,1
					elseif leixing="tejia" then
						rs.open "select * from products where tejiabook=1 and zhuangtai=0 "&orderby&"",conn,1,1
					elseif leixing="news" then
						rs.open "select * from products where newsbook=1 and zhuangtai=0 "&orderby&"",conn,1,1
					else
						rs.open "select * from products where newsbook=1 and zhuangtai=0 "&orderby&"",conn,1,1
					
					end if
					if rs.recordcount=0 then
					%>
				  
					<table width="370" border="0" cellspacing="0" cellpadding="5" align="center">
					  <tr>
						<td align=center>No goods</td>
					  </tr>
					</table>
					<%
					else
					rs.PageSize =28 '每页记录条数
					iCount=rs.RecordCount '记录总数
					iPageSize=rs.PageSize
					maxpage=rs.PageCount 
					page=request("page")
					if Not IsNumeric(page) or page="" then
						page=1
					else
						page=cint(page)
					end if
					if page<1 then
						page=1
					elseif  page>maxpage then
						page=maxpage
					end if
					rs.AbsolutePage=Page
					if page=maxpage then
						x=iCount-(maxpage-1)*iPageSize
					else
						x=iPageSize
					end if
					%>
          <table  class="sideleft"  width="100%"  border="0" cellspacing="0" cellpadding="0" align="center">
			  <tr>
                <td colspan="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=class.asp?<%=tempp%>order=PriceUp><img src="images/jk1.gif" width="60" height="21" alt="按价格从低到高"></a><a href=class.asp?<%=tempp%>order=PriceDown><img src="images/jk2.gif" width="23" height="21" alt="按价格从高到低"></a>　<a href=class.asp?<%=tempp%>order=SaleUp><img src="images/jk3.gif" width="114" height="21" alt="按销量从低到高"></a><a href=class.asp?<%=tempp%>order=SaleDown><img src="images/jk2.gif" width="23" height="21" alt="按销量从高到低"></a>　<a href=class.asp?<%=tempp%>order=ViewsUp><img src="images/jk4.gif" width="51" height="21" alt="按人气从低到高"></a><a href=class.asp?<%=tempp%>order=ViewsDown><img src="images/jk2.gif" width="23" height="21" alt="按人气从高到低"></a>　<a href=class.asp?<%=tempp%>order=Addup><img src="images/jk5.gif" width="93" height="21" alt="按上架先后排序"></a></td>
			  </tr>
			  <tr>
                <%
				ii=0
				For i=1 To x
				%>
                <td align="center"><table border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="10"></td>
                    </tr>
                    <tr>
                      <td><table border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td height="1" background="images/lineDotGray.gif"></td>
                          </tr>
                          <tr>
                            <td height="10" ></td>
                          </tr>
                          <tr>
                            <td align="center">
                              <table width="98" height="100" border="0" align="center" cellpadding="2" cellspacing="1">
                                <tbody>
                                  <tr>
                                    <td width="92" height="100" bgcolor="#ffffff" align="center"><%if rs("bookpic")="" then 
									response.write "<div align=center><a href=list.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
									else%>
                                   <a href="Products.asp?id=<%=rs("bookid")%>" target="_blank"> <img src="<%=trim(rs("bookpic"))%>" width="180" height="180" border="0" align="absmiddle" /></a>
                                   <%end if%></td>
                                  </tr>
                                </tbody>
                              </table></td>
                          </tr>
                          <tr>
                            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td height="30" align=left><%response.write "<a class=a4 href=products.asp?id="&rs("bookid")&" ><font color=#FF0000 style=""font-size:14px; font-weight:bold;"">"
									if len(trim(rs("bookname")))>12 then
									response.write left(trim(rs("bookname")),20)&".."
									else
									response.write trim(rs("bookname"))
									end if
									response.write "</font></a>"
									%></td>
                                </tr>
                                <tr>
                                  <td align=center style="background:url(images/title2bg.jpg) repeat-y center center; line-height:25px;font-family:'微软雅黑'; color:#000000;">Market value：<s><font color=#000000>$<%=trim(rs("shichangjia"))%> </font></s><br>
                                    Member price：<font color=#000000>$<%=trim(rs("huiyuanjia"))%> </font></td>
                                </tr>
                            </table></td>
                          </tr>
                      </table></td>
                    </tr>
                </table></td>
                <%rs.movenext
		     	ii=ii+1%>
                <%if ii mod 3 =0 then%>
              </tr>
              <tr>
                <%end if%>
                <% next%>
                <td colspan="2" >&nbsp;</td>
              </tr>
            </table>
			  <%
			call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
			end if
			rs.close
			set rs=nothing
			Sub PageControl(iCount,pagecount,page,table_style,font_style)
			'生成上一页下一页链接
				Dim query, a, x, temp
				action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
				query = Split(Request.ServerVariables("QUERY_STRING"), "&")
				For Each x In query
					a = Split(x, "=")
					If StrComp(a(0), "page", vbTextCompare) <> 0 Then
						temp = temp & a(0) & "=" & a(1) & "&"
					End If
				Next
				Response.Write("<table width=100% border=0 cellpadding=0 cellspacing=0 >" & vbCrLf )        
				Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR >" & vbCrLf )
				Response.Write("<TD align=center height=40>" & vbCrLf )
				Response.Write(font_style & vbCrLf )    
				if page<=1 then
					Response.Write ("Home " & vbCrLf)        
					Response.Write ("Previous " & vbCrLf)
				else        
					Response.Write("<A HREF=" & action & "?" & temp & "Page=1>Home</A> " & vbCrLf)
					Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">Previous</A> " & vbCrLf)
				end if
				if page>=pagecount then
					Response.Write ("Next " & vbCrLf)
					Response.Write ("End " & vbCrLf)            
				else
					Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">Next</A> " & vbCrLf)
					Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">End</A> " & vbCrLf)            
				end if
				Response.Write(" Ones：" & page & "/" & pageCount & "Page" &  vbCrLf)
				Response.Write(" All Found" & iCount & "Ptoduct" &  vbCrLf)
				Response.Write(" GO to" & "<INPUT CLASS=wenbenkuang TYEP=TEXT NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "Page"  & vbCrLf & "<INPUT CLASS=go-wenbenkuang type=submit value=GO>")
				Response.Write("</TD>" & vbCrLf )                
				Response.Write("</TR></form>" & vbCrLf )        
				Response.Write("</table>" & vbCrLf )        
			End Sub
			%></td>
      </tr>
	  <tr>
		<td>&nbsp;</td>
	  </tr>
    </table></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
