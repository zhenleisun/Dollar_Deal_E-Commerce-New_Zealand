<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<div align=center><font size=80 color=red><b>您没有此项目管理权限！</b></font></div>"
response.End
end if
end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td class="forumRowHighlight" height="22" bgcolor="f1f1f1"> 
      <div align="center"><font color="#000000">查看信息反馈</font></div>
    </td>
  </tr>
  <tr> 
    <td class="forumRowHighlight" height="169" align="center" valign="top" bgcolor="#FFFFFF"> 
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
      <%
set rs=server.createobject("adodb.recordset")
rs.open "select * from guestbook order by id desc",conn,3,2
'dim page
'dim e_page
'e_page=5 '每页显示留言数
'rs.pagesize=e_page
'if request.querystring("page")="" or request.querystring("page")=0 then
'page=1
'else
'page=request.querystring("page")
'rs.absolutepage=trim(request.querystring("page"))
'end if
%>

      <br> 
      <%
if rs.eof and rs.bof then
%>
      <table width="380" border="0" align="center" cellpadding="4" >
        <tr> 
          <td class="forumRowHighlight" height="40" align="center"> <p>还没有任何留言！</p></td>
        </tr>
      </table>
      <%else%>
      <%
	  	rs.PageSize =5 '每页记录条数
		iCount=rs.RecordCount '记录总数
		iPageSize=rs.PageSize
    	maxpage=rs.PageCount 
    	page=request("page")
    	per_page=rs.PageSize
    
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
      <table width="100%" border="0">
        <tr> 
          <td class="forumRowHighlight" colspan="12" height="25" align="center" bgcolor="#FFFFFF" > 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
For i=1 To x
'do while not rs.eof and e_page>0
%>
      <TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=1 bgcolor="#CCCCCC">
        <TBODY>
          <TR> 
            <td class="forumRowHighlight" bgcolor="f1f1f1"><TABLE width="100%" border=0 cellpadding="0" cellspacing="0">
                <TBODY>
                  <TR bgcolor="f1f1f1"> 
                    <td class="forumRowHighlight" width="67" height="23" align="center" bgcolor="f1f1f1"><font color="#FF0000">留言ID：<%=rs("id")%></font></TD>
                    <td class="forumRowHighlight" width="464" align="right"> 
                      <%if rs("qq")<>"" then%>
                      <a title=QQ：<%=rs("qq")%>>QQ: 
                      <%=rs("qq")%></a> 
                      <%end if%>
                      &nbsp; 
                      <%if rs("email")<>"" then%>
                     Email:
                      <a href="mailto:<%=rs("email")%>" target="_blank" title=Email：<%=rs("email")%>><%=rs("email")%></a> 
                      <%end if%>
                      <%if rs("url")<>"" then%>
                     Tel:
                      <%=rs("url")%> 
                      <%end if%>
                      &nbsp;</TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <TR> 
            <td class="forumRowHighlight" ><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                <TBODY>
                  <TR> 
                    <td class="forumRowHighlight" bgcolor="#FFFFFF"> <TABLE width="100%" border=0 style="table-layout:fixed;word-break:break-all">
                        <TBODY>
                          <TR> 
                            <td class="forumRowHighlight" width="109"><img src="../images/<%=rs("sex")%>.gif"></TD>
                            <td class="forumRowHighlight" colspan="2"> <strong><font color="<%=flzhbt%>">姓名：</font></strong><font color="<%=flzhzt%>"><%=rs("name")%></font><br> 
                              <strong><font color="<%=flzhbt%>">内容：</font></strong><font color="<%=flzhzt%>"><%=unHtml(rs("content"))%></font></TD>
                          </TR>
                          <TR> 
                            <td class="forumRowHighlight" colspan="2"><form action="reply.asp?action=save&id=<%=rs("id")%>" method="post" name="form" id="form">
                                &nbsp; [<a href="reply.asp?id=<%=rs("id")%>">回复</a>] 
                                &nbsp; [<a href="?id=<%=rs("id")%>&action=del" onclick="return confirm('确定删除此条留言吗？')">删除</a>]&nbsp;&nbsp 
                                留言状态：<input type="radio" name="Online" value="0" <%if rs("Online")="0" then%>checked<%end if%>>
                              隐藏 
                              <input type="radio" name="Online" value="1" <%if rs("Online")="1" then%>checked<%end if%>>
                                公开 
                                <input type="submit" name="Submit" value="提交">
                              </form></TD>
                            <td class="forumRowHighlight" width="250" align="right"><%=rs("time")%></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE></TD>
          </TR>
          <%if rs("reply")<>"" then%>
          <TR> 
            <td class="forumRowHighlight" bgColor=#f2f2f2> <table width="100%" border="0" style="table-layout:fixed;word-break:break-all">
                <tr> 
                  <td class="forumRowHighlight" width="10">&nbsp;</td>
                  <td class="forumRowHighlight" width="567"><font color="#FF0000">管理员回复：</font><br> <%=unHtml(rs("reply"))%></td>
                </tr>
              </table></TD>
          </TR>
          <%end if%>
        </TBODY>
      </TABLE>
      <br> 
      <%
'e_page=e_page-1
'rs.movenext
'loop
		RS.MoveNext
next
%>
      <table width="100%" border="0">
        <tr> 
          <td class="forumRowHighlight" > 
            <%
					call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>",per_page)
				  %>
          </td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
end if
%>
    </td>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html><%
Sub PageControl(iCount,pagecount,page,table_style,font_style,per_page)
'生成上一页下一页链接
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
	
    temp=""

    Response.Write("<table " & Table_style & ">" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<td  align=right>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
        
    if page<=1 then
        Response.Write ("首页 " & vbCrLf)        
        Response.Write ("上页 " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>首页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">上页</A> " & vbCrLf)
    end if

    if page>=pagecount then
        Response.Write ("下页 " & vbCrLf)
        Response.Write ("尾页 " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">下页</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">尾页</A> " & vbCrLf)            
    end if

    Response.Write(" 页次：" & page & "/" & pageCount & "页" &  vbCrLf)
    Response.Write(" 共有" & iCount & "条/每页"&per_page&"条" &  vbCrLf)
    Response.Write(" 转到" & "<INPUT TYEP=TEXT NAME=page SIZE=4 Maxlength=8 VALUE=" & page & ">" & "页"  & vbCrLf & "<INPUT type=submit style=""font-size: 9pt"" value=GO class=b2>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>
<%
if request("action")="del" then
set rs=server.createobject("adodb.recordset")
conn.execute "delete from guestbook where id="&trim(request.querystring("id"))
response.redirect "viewreturn.asp"
end if 
%>