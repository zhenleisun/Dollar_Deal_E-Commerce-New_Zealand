<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<body>

<%dim selectm,selectkey,selectbookid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if
selectbookid=request("selectbookid")
if selectbookid<>"" then
conn.execute "delete FROM inform where newsid in ("&selectbookid&")"
response.Redirect "manageinforms.asp"
response.End
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
</tr>
<tr> 
<form name="form1" method="post" action="">
<td height="100"> 
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
			set rs=server.CreateObject("adodb.recordset")
			select case selectm
			case ""
            rs.open "select newsid,title,updatetime FROM inform order by updatetime desc",conn,1,1
		    case "0"
			rs.open "select newsid,title,updatetime FROM inform order by updatetime desc",conn,1,1
		    case "newsid"
			rs.open "select newsid,title,updatetime FROM inform where newsid="&selectkey&" order by updatetime desc",conn,1,1
			case "title"
			rs.open "select newsid,title,updatetime FROM inform where title like '%"&selectkey&"%' order by updatetime desc",conn,1,1
			case "content"
			rs.open "select newsid,title,updatetime FROM inform where content like '%"&selectkey&"%' order by updatetime desc",conn,1,1
		  end select
		   	if err.number<>0 then
				response.write "���ݿ���������"
				end if
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> ���ݿ��������ݣ�</p>"
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
            			showpage totalput,MaxPerPage,"manageinforms.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"manageinforms.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"manageinforms.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0%>
                <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      				    <tr> 
      				      <td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�̳�ר�����</font></b></td>
			      </tr>
                                  <tr > 
                                    <td width="10%" align="center" bgcolor="fbc2c2">���</td>
                                    <td width="50%" align="center" bgcolor="fbc2c2">���±���</td>
                                    <td width="30%" align="center" bgcolor="fbc2c2">����ʱ��</td>
                                    <td width="10%" align="center" bgcolor="fbc2c2">ѡ ��</td>
                                  </tr>
                                  <%
		  do while not rs.eof%>
                                  <tr > 
                                    <td> 
                                      <div align="center"><a href=editbook.asp?id=<%=rs("newsid")%>> 
                                        </a><%=rs("newsid")%></div>
                                    </td>
                                    <td width="50%" STYLE='PADDING-LEFT: 10px'> 
                                      <a href=editinforms.asp?id=<%=rs("newsid")%>> 
                                      <% if len(replace(trim(rs("title")),"<br>",""))>16 then
			response.write left(replace(trim(rs("title")),"<br>",""),14)&"..."
			else
			response.write replace(trim(rs("title")),"<br>","")
			end if%>
                                      </a></td>
                                    <td align="center"><%=rs("updatetime")%></td>
                                    <td align="center">
									<input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("newsid")%>">
                                    </td>
                                  </tr>
                                  <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                                  <tr > 
                                    <td height="30" colspan="4" align="right">
									ȫѡ
									<input type="checkbox" name="checkbox" value="Check All" onClick="mm()">
									<input type="submit" name="Submit" value="ɾ ��" onClick="return test();">
									&nbsp;&nbsp;
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
				Response.Write "<form method=Post action="&filename&"?selectm="&selectm&"&selectkey="&selectkey&" >"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>��һҳ</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>βҳ</a>"  
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"ר�� " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok' ></form>"  
				End Function  
			%>
                                <table width="12" height="7" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height=7></td>
                                  </tr>
                                </table>
      </td>
                            </form>
                          </tr>
                        </table>
					<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
      				    <tr> 
      				      <td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">ר�����²�ѯ</font></b></td>
     				     </tr>
                          <tr> 
                            <form name="form2" method="post" action="manageinforms.asp">
                              <td height="50" > 
                                <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="35%" align="center">
									<input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؼ���">
                                    </td>
                                    <td width="45%" align="center">
                                        <select name="selectm" id="selectm">
                                          <option value="0">ȫ������</option>
                                          <option value="newsid">���������</option>
                                          <option value="title">�����±���</option>
                                          <option value="content">����������</option>
                                        </select>
                                    </td>
                                    <td width="20%" height="30" align="center">
									<input type="submit" name="Submit2" value="�� ѯ">
                                    </td>
                                  </tr>
								  <tr><td colspan=3 height=40>������ѯ��������Ʒ���⣬����Ҫ����ؼ��֣����ܲ�ѯ<br>��������Ʒ�������ѯʱ���ؼ�����ֻ���������֡�</td></tr>
                                </table>
                              </td>
                            </form>
                          </tr>
                        </table>
<!--#include file="foot.asp"-->
</body>
</html>
<script>
function test()
{
  if(!confirm('ȷ��ɾ����')) return false;
}
</script>
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
