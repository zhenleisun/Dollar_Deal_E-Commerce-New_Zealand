<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<div align=center><font size=80 color=red><b>��û�д���Ŀ����Ȩ�ޣ�</b></font></div>"
response.End
end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<%dim action
action=request.QueryString("action")
if InStr(action,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<form name="form1" method="post" action="">
<td> 
                                <%'��ʼ��ҳ
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
		  select case action
		  case "no"
		  rs.open "select products.bookname,products.bookid,review.pinglunid,review.pingluncontent,review.pingluntitle,review.pinglundate from review,products where products.bookid=review.bookid and review.shenhe=0 order by review.bookid desc,review.pinglunid desc",conn,1,1
		  case "hf"
		  rs.open "select products.bookname,products.bookid,review.pinglunid,review.pingluncontent,review.pingluntitle,review.pinglundate from review,products where products.bookid=review.bookid and (review.huifu is null) order by review.bookid desc,review.pinglunid desc",conn,1,1
		  case "all"
		  rs.open "select products.bookname,products.bookid,review.pinglunid,review.pinglunname,review.pingluntitle,review.pingluncontent,review.pinglundate from review,products where products.bookid=review.bookid order by review.bookid desc,review.pinglunid desc",conn,1,1
		  case "yes"
		  		  rs.open "select products.bookname,products.bookid,review.pinglunid,review.pingluncontent,review.pingluntitle,review.pinglundate from review,products where products.bookid=review.bookid and review.shenhe=1 order by review.bookid desc,review.pinglunid desc",conn,1,1
		  end select
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
				if rs.eof And rs.bof then
       			Response.Write "<p align='center' class='contents'> û���κ����ۣ�</p>"
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
            			showpage totalput,MaxPerPage,"managereview.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"managereview.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"managereview.asp"
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
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><font color="#ffffff">
�������� <a href="managereview.asp?action=all"><b><font color="#ffffff">��������</font></b></a>
�������� <a href="managereview.asp?action=hf"><b><font color="#ffffff">δ�ظ�������</font></b></a>
�������� <a href="managereview.asp?action=no"><b><font color="#ffffff">δ��˵�����</font></b></a>
�������� <a href="managereview.asp?action=yes"><b><font color="#ffffff">����˵�����</font></b></a></font>
</td></tr>
<tr > 
<td width="30%" align="center" bgcolor="fbc2c2">������Ʒ����</td>
<td width="30%" align="center" bgcolor="fbc2c2">���۱���</td>
<td width="30%" align="center" bgcolor="fbc2c2">����ʱ��</td>
<td width="10%" align="center" bgcolor="fbc2c2">�� ��</td>
</tr>
	<%  
		do while not rs.eof%>
<tr > 
<td align="center">
			<%if len(rs("bookname"))>14 then
			response.write "<a href=../products.asp?id="&rs("bookid")&" >"&left(trim(rs("bookname")),12)&"...</a>"
			else
			response.write "<a href=../products.asp?id="&rs("bookid")&" >"&trim(rs("bookname"))&"</a>"
			end if
			%>
</td>
<td align="center">
			<% if len(rs("pingluntitle"))>12 then
			response.write "<a href=""javascript:;"" onClick=""javascript:window.open('review.asp?id="&rs("pinglunid")&"','pinglun','width=400,height=350,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=no');return false;"" title="&trim(rs("pingluntitle"))&">"&left(trim(rs("pingluntitle")),10)&"...</a>"
			else
			response.write "<a href=""javascript:;"" onClick=""javascript:window.open('review.asp?id="&rs("pinglunid")&"','pinglun','width=400,height=350,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=no');return false;"" title="&trim(rs("pingluntitle"))&">"&trim(rs("pingluntitle"))&"</a>"
			end if%>
</td>
<td align="center"><%=rs("pinglundate")%></td>
<td align="center"><input name="shenhe" type="checkbox" id="shenhe" value="<%=rs("pinglunid")%>"></td>
</tr>
			<%i=i+1
		  if i>=MaxPerPage then Exit Do
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  %>
<tr > 
<td height="30" colspan="4" align="center">
<%if action="no" then%>
<input type="submit" name="Submit" value="ͨ�����" onClick="this.form.action='savereview.asp?action=shenhe';this.form.submit()">
<%end if%>
&nbsp; 
<input type="button" name="Submit2" value="ɾ ��" onClick="this.form.action='savereview.asp?action=del';this.form.submit()">
&nbsp;&nbsp;ȫѡ 
<input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
</td>
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
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&action="&action&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&action="&action&" class='contents'>��һҳ</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&action="&action&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&action="&action&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"����¼ " 
					Response.Write "<font class='contents'>" 
					Response.Write "</form>"  
				End Function  
			%>
</td>
</form>
</tr>
</table>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�� �� �� ��</font></b> </td>
</tr>
<tr> 
<td height="50" > 
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td height="20" class=pad>ɾ��һ��ǰδ��˵����ۡ�
<input type="submit" name="Submit4" value="ȷ ��" onClick="if(confirm('��ȷ������������?')) location.href='savereview.asp?action=delzhou';else return;">
</td>
</tr>
<tr> 
<td height="16" class=pad>ɾ������δ��˵����ۡ���
<input type="submit" name="Submit4" value="ȷ ��" onClick="if(confirm('��ȷ������������?')) location.href='savereview.asp?action=delall';else return;">
</td>
</tr>
</table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
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
