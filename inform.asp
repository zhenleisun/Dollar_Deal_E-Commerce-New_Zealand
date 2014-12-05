<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--行业资讯</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<meta name="keywords" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="include/head.asp"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="39" valign="top"><table width="100%" height="184" border="0" cellpadding="0" cellspacing="0" >
      <tr>
        <td height="30" align="right" valign="top"><img src="images/body/jiao.gif" width="15" height="17"></td>
      </tr>
      <tr>
        <td height="90"><div align="right"><img src="images/ttts.gif" width="26" height="110" /></div></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table></td>
    <td width="190"valign="top" style="BORDER-bottom:#FF67A0 1px solid;BORDER-left:#FF67A0 1px solid; BORDER-right:#FF67A0 1px solid;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
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
    <td valign="top"><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
      <tr>
        <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> 行业资讯</td>
      </tr>
      <tr>
        <td><br>
          <img src="images/mingle/inform.gif" width="643" height="59"></td>
      </tr>
      <tr>
        <td width="50%" align="center" valign="top" bordercolor="#FFFFFF" bgcolor="#FFFFFF" ><br>
          <table width="580" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/mingle/informtop.gif" width="580" height="35"></td>
              </tr>
          </table>
          <br>
          <%set rs=server.createobject("adodb.recordset")
		rs.open "select * FROM inform  order by newsid desc",conn,1,1
		if rs.recordcount=0 then 
		%>
            <table width="100%" border="0" cellspacing="0" cellpadding="5" align="center">
              <tr>
                <td align=center>暂无专题</td>
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
          <table width="580" border="0" align="center" cellpadding="2" cellspacing="1">
              <%
			For i=1 To x
			%>
              <tr>
                <td width="6%" height=24 align=center bgcolor="#FFFFFF" ><img src="images/body/dian3.gif"></td>
                <td width="66%" height=24 bgcolor="#FFFFFF" > <a href=informs.asp?id=<%=rs("newsid")%> >  <%=replace(trim(rs("title")),"<br>","")%></a>&nbsp<font color=red>(<%=rs("viewcountent")%>)</font></td>
                <td width="28%" height=24 align=center bgcolor="#FFFFFF" ><%=year(rs("updatetime"))&"."&month(rs("updatetime"))&"."&day(rs("updatetime"))%></td>
              </tr>
              <tr>
                <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
              </tr>
              <%
			rs.movenext
		    next
			%>
            </table>
            <br>
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
    Response.Write("<table width=90% border=0 cellpadding=0 cellspacing=0 >" & vbCrLf )        
    Response.Write("<form method=get onsubmit=""document.location = '" & action & "?" & temp & "Page='+ this.page.value;return false;""><TR>" & vbCrLf )
    Response.Write("<TD align=center height=35>" & vbCrLf )
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
    Response.Write(" 共有" & iCount & "条专题" &  vbCrLf)
    Response.Write(" 转到" & "<INPUT TYEP=TEXT CLASS=wenbenkuang NAME=page SIZE=2 Maxlength=5 VALUE=" & page & ">" & "页"  & vbCrLf & "<INPUT CLASS=go-wenbenkuang type=submit value=GO>")
    Response.Write("</TD>" & vbCrLf )                
    Response.Write("</TR></form>" & vbCrLf )        
    Response.Write("</table>" & vbCrLf )        
End Sub
%>        </td>
      </tr>
    </table></td>
    <td width="68"style="BORDER-left:#FF67A0 1px solid;">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
