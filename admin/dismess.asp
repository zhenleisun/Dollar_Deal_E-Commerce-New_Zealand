<!--#include file="conn.asp"-->
<%
if session("admin")="" then
conn.close
set conn = nothing
response.Write "<script language='javascript'>alert('���ȵ�¼��');history.go(-1);</script>"
response.End
else
if request.cookies("sundxshop")("admin")="" then
conn.close
set conn = nothing
response.Write "<script language='javascript'>alert('���ȵ�¼��');history.go(-1);</script>"
response.End
end if
end if

if request.form("action")="del" then
if session("rank")>1 then
conn.close
set conn = nothing
response.Write "<script language='javascript'>alert('����Ȩɾ�����ԣ�');history.go(-1);</script>"
response.End
end if
conn.execute "delete from mess where messid in ("&request.form("selectdel")&")"
response.Redirect  "dismess.asp"
end if
%>
<title>�鿴���Է���</title>
<TABLE WIDTH="98%" BORDER="0" ALIGN="center" CELLPADDING="0" CELLSPACING="0" class="table-zuoyou">
  <TR ALIGN="right"> 
    <TD height="30" class="table-youxia" width="55%"><b>�鿴��Ϣ����</b></TD>
    <TD class="table-xia">ѡ��鿴���� 
            <SELECT NAME="select" ONCHANGE="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" ><BASE TARGET=Right> 
              <OPTION selected>ѡ����Ϣ����</OPTION>
              <OPTION VALUE="dismess.asp?lx=0" >��������</OPTION>
              <OPTION VALUE="dismess.asp?lx=1" >��ͨ����</OPTION>
              <OPTION VALUE="dismess.asp?lx=2" >�������</OPTION>
              <OPTION VALUE="dismess.asp?lx=3" >ȱ���Ǽ�</OPTION>
              <OPTION VALUE="dismess.asp?lx=4" >��������</OPTION>
              <OPTION VALUE="dismess.asp?lx=5" >��ƷͶ��</OPTION>
              <OPTION VALUE="dismess.asp?lx=6" >����Ͷ��</OPTION>
            </SELECT>&nbsp;</TD>
  </TR></table>
      <%
sub leixing()
select case rs("messtype")
case "1"
response.write "��ͨ����"
case "2"
response.write "�������"
case "3"
response.write "ȱ���Ǽ�"
case "4"
response.write "��������"
case "5"
response.write "��ƷͶ��"
case "6"
response.write "����Ͷ��"
end select
end sub
%> <%
				Const MaxPerPage=10 
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages
   				dim j
   				dim sql
    				if Not isempty(SafeRequest("page",1)) then
      				currentPage=Cint(SafeRequest("page",1))
   				else
      				currentPage=1
   				end if 
		  dim lx
	  lx=SafeRequest("lx",1)
	  set rs=server.CreateObject("adodb.recordset")
	  if lx="" or lx=0 then
	  rs.open "select * from mess order by messdtm desc",conn,1,1
	  else
	  rs.open "select * from mess where messtype="&lx&" order by messdtm desc",conn,1,1
	  end if
		  
				if err.number<>0 then
				response.write "���ݿ�����ʱ���������"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> �Բ�����ʱ��û���κ����ԣ�</p>"
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
            			showpage totalput,MaxPerPage,"dismess.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"dismess.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"dismess.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0

			%>
	
	<TABLE WIDTH="98%" BORDER="0" ALIGN="center" class=table-zuoyou CELLPADDING="0" CELLSPACING="0" BGCOLOR="#FFFFFF">
    <TR BGCOLOR="#FFFFFF" height=25 align="center"> 
      <TD WIDTH="10%" class=table-youxia nowrap>��������</TD>
      <TD WIDTH="40%" class=table-youxia>���Ա���</TD>
      <TD WIDTH="20%" class=table-youxia>�����û�</TD>
      <TD WIDTH="20%" class=table-youxia>����ʱ��</TD>
      <TD width="40" class=table-youxia>ѡ��</TD>
      <TD width="40" class=table-xia nowrap>�� ��</TD>
    </TR>
	<form name="form1" method="post" action="dismess.asp?action=del">
    <% do while not rs.eof %>
    <TR BGCOLOR="#FFFFFF" height=20 align="center"> 
      <TD class=table-youxia> <% call leixing() %> </TD>
      <TD class=table-youxia align="left">&nbsp; <a href=# onClick="javascript:window.open('../chkmess.asp?id=<% = rs("messid") %>','','width=500,height=300,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"> 
        <%=trim(rs("messsubject"))%></a></TD>
      <TD class=table-youxia><a href=mailto:<% = trim(rs("messemail"))%>> 
        <% = trim(rs("messusername")) %>
        </a></TD>
      <TD class=table-youxia>&nbsp;<%= rs("messdtm")%></TD>
      <TD class=table-youxia><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("messid")%>></TD>
      <TD class=table-xia><a href=replymess.asp?id=<%=rs("messid")%>>�ظ�</a></TD>
    </TR>
	<%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
	  loop 
	  rs.close
	  set rs=nothing%>
    <TR BGCOLOR="#FFFFFF" height=20 align="center"> 
      <TD colspan="6" class=table-youxia><br>
		<input name="action" type="hidden" value="del">
	  <input type="submit" name="Submit" value="ɾ����ѡ����">
                ȫѡ<input type="checkbox" name="checkbox" value="Check All" onClick="mm()"></TD>
    </TR></form>
  </TABLE>
	  
	  <TABLE WIDTH="98%" BORDER="0" ALIGN="center" class=table-zuoyou CELLPADDING="0" CELLSPACING="0">
        <TR align="center">
		<td class=table-xia>
      <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				
				Response.Write "<form method=Post action="&filename&"?lx="&lx&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>�� ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&lx="&lx&" class='contents'>�� ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&lx="&lx&" class='contents'>��һҳ</a> "  
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ ĩ ҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&lx="&lx&"&class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&lx="&lx&"&class='contents'>ĩ ҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"����Ϣ " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='��ת' name='cndok'></form>"  
				End Function  
			%> </td></tr>
			</table>
<!--#include file="footer.asp"-->
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