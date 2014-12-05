<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
action=request("action")
id=request("id")
	
if action="del"  and id="" then
set rs=server.CreateObject("adodb.recordset")
	rs.open "delete  * from words ",conn,1,3
	elseif action="del"  and id<>"" then
	set rs=server.CreateObject("adodb.recordset")
	rs.open "delete  * from words where id="&id,conn,1,3
	response.redirect  "searchword.asp"
	end if 
%>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<form name="form1" method="post" action="">
<td> 
                                <%'开始分页
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

		  rs.open "select * from words order by times desc ",conn,1,1
	
				if err.number<>0 then
				response.write "数据库中无数据"
				end if
				if rs.eof And rs.bof then
       			Response.Write "<br><br><p align='center' class='contents'> 没有搜索关键词信息！</p>"
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
            			showpage totalput,MaxPerPage,"searchword.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"searchword.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"searchword.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0
			response.write "<table width=12 height=7 border=0 cellpadding=0 cellspacing=0><tr><td height=7></td></tr></table>"
			%>
        <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
          <tr> 
            <td colspan="6" align="center" background="../images/admin_bg_1.gif"><font color="#ffffff"> 
              　<b>商品<font color="#ffffff">搜索关键词</font></b></font> </td>
          </tr>
          <tr > 
            <td width="29%" align="center" bgcolor="fbc2c2">搜索关键字</td>
            <td width="12%" align="center" bgcolor="fbc2c2">搜索次数</td>
            <td width="22%" align="center" bgcolor="fbc2c2">搜索时间</td>
            <td colspan="2" align="center" bgcolor="fbc2c2">搜索用户IP</td>
            <td width="13%" align="center" bgcolor="fbc2c2">删除</td>
          </tr>
          <%  
		do while not rs.eof%>
          <tr > 
            <td align="center"> <%=rs("keyword")%> </td>
            <td align="center"> <%=rs("times")%> </td>
            <td align="center"><%=rs("date")%></td>
            <td width="4%" align="left">&nbsp;</td>
            <td width="20%" align="left"><%=rs("ip")%></td>
            <td align="left"><div align="center"><a href="?id=<%=rs("id")%>&action=del">删除</a></div></td>
          </tr>
          <%i=i+1
		  if i>=MaxPerPage then Exit Do
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  %>
          <tr > 
            <td height="30" colspan="6" align="center"> &nbsp; <input type="button" name="Submit2" value="全部清空" onClick="this.form.action='searchword.asp?action=del';this.form.submit()"> 
              &nbsp;&nbsp;</td>
          </tr>
        </table>
				<%  
				End Sub   
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				Response.Write "<form method=Post action="&filename&"?action="&action&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&action="&action&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&action="&action&" class='contents'>上一页</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&action="&action&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&action="&action&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"条记录 " 
					Response.Write "<font class='contents'>" 
					Response.Write "</form>"  
				End Function  
			%>
</td>
</form>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
