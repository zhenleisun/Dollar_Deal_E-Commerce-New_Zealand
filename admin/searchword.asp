<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
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

		  rs.open "select * from words order by times desc ",conn,1,1
	
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
				if rs.eof And rs.bof then
       			Response.Write "<br><br><p align='center' class='contents'> û�������ؼ�����Ϣ��</p>"
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
              ��<b>��Ʒ<font color="#ffffff">�����ؼ���</font></b></font> </td>
          </tr>
          <tr > 
            <td width="29%" align="center" bgcolor="fbc2c2">�����ؼ���</td>
            <td width="12%" align="center" bgcolor="fbc2c2">��������</td>
            <td width="22%" align="center" bgcolor="fbc2c2">����ʱ��</td>
            <td colspan="2" align="center" bgcolor="fbc2c2">�����û�IP</td>
            <td width="13%" align="center" bgcolor="fbc2c2">ɾ��</td>
          </tr>
          <%  
		do while not rs.eof%>
          <tr > 
            <td align="center"> <%=rs("keyword")%> </td>
            <td align="center"> <%=rs("times")%> </td>
            <td align="center"><%=rs("date")%></td>
            <td width="4%" align="left">&nbsp;</td>
            <td width="20%" align="left"><%=rs("ip")%></td>
            <td align="left"><div align="center"><a href="?id=<%=rs("id")%>&action=del">ɾ��</a></div></td>
          </tr>
          <%i=i+1
		  if i>=MaxPerPage then Exit Do
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  %>
          <tr > 
            <td height="30" colspan="6" align="center"> &nbsp; <input type="button" name="Submit2" value="ȫ�����" onClick="this.form.action='searchword.asp?action=del';this.form.submit()"> 
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
<!--#include file="foot.asp"-->
</body>
</html>
