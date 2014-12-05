<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--Mall news</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="#">
<meta name="keywords" content="#">

<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background:url(images/ds_07.jpg); text-align:center;">
<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" align="left"valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12"></td>
      </tr>
      <tr>
        <td><!--#include file="searc.asp"--></td>
      </tr>
	  <tr>
        <td background="images/leftbg.jpg"><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg"></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
      <tr>
        <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Mall news</td>
      </tr>
      <tr>
        <td valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><br>
          <table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><img src="images/mingle/informtop.gif" width="700" height="35"></td>
            </tr>
          </table>
          <br>
          <%set rs=server.createobject("adodb.recordset")
		rs.open "select * from news  order by adddate desc",conn,1,1
		if rs.recordcount=0 then 
		%>
            <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
              <tr>
                <td align=center>No news</td>
              </tr>
            </table>
            <%
		else
	  		rs.PageSize =20 
			iCount=rs.RecordCount 
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
            <table width="700" border="0" align="center" cellpadding="2" cellspacing="1">
              <%
			For i=1 To x
			%>
              <tr>
                <td width="7%" height=24 align=center bgcolor="#FFFFFF" ><img src="images/body/dian3.gif"></td>
                <td width="77%"  height=24 bgcolor="#FFFFFF" ><a href="trends.asp?id=<%=rs("newsid")%>"><%=trim(rs("newsname"))%></a>&nbsp;<font color=red>(<%=rs("viewcount")%>)</font></td>
                <td width="16%" height=24 align=left bgcolor="#FFFFFF" ><%=year(rs("adddate"))&"."&month(rs("adddate"))&"."&day(rs("adddate"))%></td>
              </tr>
              <tr>
                <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
              </tr>
              <%
			rs.movenext
		    next
			%>
            </table>
            <%
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
Sub PageControl(iCount,pagecount,page,table_style,font_style)
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    Response.Write("<table width=90% border=0 cellpadding=0 cellspacing=0  align=center >" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=center height=35>" & vbCrLf )
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
    Response.Write(" Found "& iCount &" news" &  vbCrLf)
    Response.Write(" Go to" & "<INPUT TYEP=TEXT CLASS=wenbenkuang NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "Page"  & vbCrLf & "<INPUT CLASS=go-wenbenkuang type=submit value=GO>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>        </td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
