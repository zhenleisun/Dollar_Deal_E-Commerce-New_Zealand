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
%>
<%
 if request.QueryString("action")="del" then
conn.execute "delete from dyemail where id in ("&request("selectdel")&")"
response.Redirect  "dmail.asp"
end if %>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr > 
    <th height=25  background="../images/admin_bg_1.gif" colspan=6 class="tableHeaderText">邮件订阅管理</th>
  </tr>
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
		 set rs=server.createobject("adodb.recordset")
	rs.open "select * from dyemail order by id desc",conn,1,1
		  
				if err.number<>0 then
				response.write "数据库中无数据"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 您还没有添加邮件！</p>"
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
            			showpage totalput,MaxPerPage,"dmail.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"dmail.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"dmail.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0

			%>
  <form name="form1" method="post" action="dmail.asp?action=del">
    <tr> 
      <td class="forumRowHighlight"> <div align="center">ID号</div></td>
      <td class="forumRowHighlight"> <div align="center">订阅邮址</div></td>
      <td class="forumRowHighlight"> <div align="center">订阅时间</div></td>
      <td class="forumRowHighlight"> <div align="center">选 择</div></td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td class="forumRowHighlight"><div align="center"><%=rs("id")%></div></td>
      <td class="forumRowHighlight"> <div align="center"><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></div></td>
      <td class="forumRowHighlight"> <div align="center"><%=rs("dingdate")%></div></td>
      <td class="forumRowHighlight"> <div align="center"> 
          <input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("id")%>>
        </div></td>
    </tr>
    <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
    <tr> 
      <td height="30" colspan="4" class="forumRowHighlight"> <div align="center"> 
          <input   type="submit" name="Submit" value="删除所选邮件">
          　　全选 
          <input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
        </div></td>
    </tr>
  </form>
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
				
				Response.Write "<form method=Post action="&filename&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1 class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&" class='contents'>上一页</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"条邮件 " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='button' value='GO' name='cndok'></form>"  
				End Function  
%>
<br>
</body>
</html>
<script language=javascript>
function mm()
{
   var a = document.getElementsByTagName("input");
   if(a[0].checked==true){
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = false;
   }
   else
   {
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = true;
   }
}
</script>