<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<%
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
		  dim namekey,checkbox,action
		  action=request.QueryString("action")
		  checkbox=request("checkbox")
		  namekey=request("namekey")
		  if InStr(namekey,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
		  if namekey="" then namekey=request.QueryString("namekey")
		  if checkbox="" then checkbox=request.querystring("checkbox")
		 '//
		 set rs=server.CreateObject("adodb.recordset")
		 if namekey="" then
		  select case action
		  case "all"
		 	rs.open "select username,userid,userzhenshiname,logins,adddate from [user]  where niming=0 order by adddate desc",conn,1,1
   		  case "niming"
		 	rs.open "select username,userid,userzhenshiname,logins,adddate from [user]  where niming=1 order by adddate desc",conn,1,1
		  case "huiyuan"
		 	rs.open "select username,userid,userzhenshiname,logins,adddate from [user] where reglx=1  order by adddate desc",conn,1,1
		  case "vip"
		 	rs.open "select username,userid,userzhenshiname,logins,adddate from [user] where reglx=2  order by adddate desc",conn,1,1
		  end select
		  else
		  if checkbox=1 then
		  rs.open "select username,userid,userzhenshiname,logins,adddate from [user] where username like '%"&namekey&"%'  order by adddate desc ",conn,1,1
		  else
		  rs.open "select username,userid,userzhenshiname,logins,adddate from [user] where username='"&namekey&"'  order by adddate desc ",conn,1,1
		  end if
		  end if
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> �Բ���û���ҵ����û���</p>"
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
            			showpage totalput,MaxPerPage,"manageuser.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"manageuser.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"manageuser.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0
			%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<tr align="center"> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�û�����</font></b></td>
</tr>
<tr> 
<form name="form1" method="post" action="saveuser.asp?action=del">
<td height="100" valign="top">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" >
<tr  align="center">
<td width="25%" bgcolor="fbc2c2"> �û���</td>
<td width="25%" height="20" bgcolor="fbc2c2"> ��ʵ����</td>
<td width="25%" bgcolor="fbc2c2"> ע��ʱ��</td>
<td width="15%" bgcolor="fbc2c2"> ��½����</td>
<td width="10%" bgcolor="fbc2c2"> ѡ ��</td>
</tr>
<%do while not rs.eof%>
<tr >
<td align="center" style="PADDING-LEFT: 10px"><a href=userlist.asp?id=<%=rs("userid")%>><%=trim(rs("username"))%></a></td>
<td align="center" style="PADDING-LEFT: 10px"><%=trim(rs("userzhenshiname"))%></td>
<td align="center" style="PADDING-LEFT: 10px"><%=rs("adddate")%></td>
<td align="center"><%=rs("logins")%> ��</td>
<td align="center"><input name="userid" type="checkbox" id="userid" value="<%=rs("userid")%>"></td>
</tr>
		<%i=i+1
		if i>=MaxPerPage then Exit Do
		rs.movenext
		loop%>
</table>
<p align="center"> 
<input type="submit" name="Submit" value="ɾ����ѡ�û�" onClick="return confirm('��ȷ��Ҫɾ�����û���')">&nbsp;&nbsp;
<input type="checkbox" name="checkbox2" value="Check All" onClick="mm()">
ȫѡ</p></td>
</form>
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
				Response.Write "<form method=Post action="&filename&"?action="&action&"&checkbox="&checkbox&"&namekey="&namekey&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?action="&action&"&page=1&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?action="&action&"&page="&CurrentPage-1&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>��һҳ</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?action="&action&"&page="&(CurrentPage+1)&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?action="&action&"&page="&n&"&checkbox="&checkbox&"&namekey="&namekey&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"��ע���û� " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok'></form>"  
				End Function  
			%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�� �� �� ��</font></b></td>
</tr>
<tr> 
<td height="50" > 
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<form name="form2" method="post" action="manageuser.asp?action=select">
<td align="center">
���û�������: 
<input name="namekey" type="text" id="namekey" size="12">&nbsp;
<input name="checkbox" type="checkbox" id="checkbox" value="1" checked>
ģ����ѯ 
<input type="submit" name="Submit2" value=" ��ʼ��ѯ ">
</div>
</td>
</form>
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
