
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
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
  <%dim selectm,selectkey,selectbookid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
'//ɾ����Ʒ
if selectm="" then
selectm=request.QueryString("selectm")
end if
selectbookid=request("selectbookid")
if selectbookid<>"" then
conn.execute "delete from award where bookid in ("&selectbookid&")"
response.Redirect "ManageAward.asp"
response.End
end if
%>
<form name="form1" method="post" action="">
  <table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
    <tr> 
      <th class="tableHeaderText" colspan=4>�鿴���޸Ľ�Ʒ</th>
    </tr>
	<div align="center">
    <%'��ʼ��ҳ
				Const MaxPerPage=10
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
            rs.open "select * from award",conn,1,1
		    case "all"
			rs.open "select bookid,bookname,adddate from award order by adddate desc",conn,1,1
		    case "bookid"
			rs.open "select bookid,bookname,adddate from award where bookid="&selectkey&"",conn,1,1
			case "bookname"
			rs.open "select bookid,bookname,adddate from award where bookname like '%"&selectkey&"%'",conn,1,1
			case "bookcontent"
			rs.open "select bookid,bookname,adddate from award where bookcontent like '%"&selectkey&"%'",conn,1,1
		  end select
		   	if err.number<>0 then
				response.write "���ݿ���������"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> �Բ���,û�������ҵĽ�Ʒ��</p>"
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
            			showpage totalput,MaxPerPage,"ManageAward.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"ManageAward.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"ManageAward.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0%>
				</div>
    <tr bgcolor="#E8F1FF"> 
      <td width="10%" class="forumRowHighlight"> <div align="center">���</div></td>
      <td width="50%" class="forumRowHighlight"> <div align="center">��Ʒ����</div></td>
      <td width="30%" class="forumRowHighlight"> <div align="center">����ʱ��</div></td>
      <td width="10%" class="forumRowHighlight"> <div align="center">ѡ ��</div></td>
    </tr>
    <%
		  do while not rs.eof%>
    <tr> 
      <td class="forumRowHighlight"> <div align="center"><%=rs("bookid")%></div></td>
      <td class="forumRowHighlight"> <a href=EditAward.asp?id=<%=rs("bookid")%>> 
        <% if len(trim(rs("bookname")))>16 then
			response.write left(trim(rs("bookname")),14)&"..."
			else
			response.write trim(rs("bookname"))
			end if%>
        </a></td>
      <td class="forumRowHighlight"> <div align="center"><%=rs("adddate")%></div></td>
      <td class="forumRowHighlight"> <div align="center"> 
          <input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("bookid")%>">
        </div></td>
    </tr>
    <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
    <tr> 
      <td class="forumRowHighlight" colspan="4"> <div align="right"> 
          <input class="button" type="submit" name="Submit" value="ɾ ��" onClick="return test();">
        </div></td>
    </tr>
	</table>
	<br>
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr align="center"> 
      <td> 
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
					Response.Write "<font class='contents'> ����"&totalnumber&"����Ʒ " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='button' value='GO' name='cndok' ></form>"  
				End Function  
			%> </td>
    </tr>
  </table>
</form>

<form name="form2" method="post" action="ManageAward.asp">
  <table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
    <tr> 
      <th class="tableHeaderText" colspan=3>��Ʒ��Ѷ</th>
    </tr>
    <tr> 
      <td width="30%" class="forumRowHighlight">
	  <div align="center"> <input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="������ؼ���">
        </div></td>
      <td width="70%" class="forumRowHighlight"> <div align="center"> 
          <select name="selectm" id="selectm">
            <option value="all" selected>ȫ����Ʒ</option>
            <option value="bookid">����Ʒ���</option>
            <option value="bookname" >����Ʒ����</option>
            <option value="bookcontent">����Ʒ˵��</option>
          </select>
        </div></td>
      <td class="forumRowHighlight"> <div align="center"><input class=button type="submit" name="Submit2" value="�� Ѷ"></div></td>
    </tr>
  </table>
</form>
</body>
</html>
<script>
function test()
{
  if(!confirm('ȷ��ɾ����')) return false;
}
</script>