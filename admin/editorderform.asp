<!--#include file="conn.asp"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%dim zhuangtai,namekey
namekey=trim(request("namekey"))
zhuangtai=trim(request("zhuangtai"))
if zhuangtai="" then zhuangtai=request.QueryString("zhuangtai")
if namekey="" then namekey=request.querystring("namekey")
%>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
   <tr> 
	<th class="tableHeaderText" colspan=6>������ѯ</th>	
   </tr>
  <tr> 
    <form name="form1" method="post" action="editorderform.asp">
      <td class="forumRowHighlight"> <div align="center">���µ��û���ѯ 
          <input name="namekey" type="text" id="namekey" value="�������û���" size="14" onFocus="this.value=''">
          &nbsp; 
          <select name="zhuangtai" id="zhuangtai">
            <option value="0" selected>--ѡ���ѯ״̬--</option>
            <option value="0" >ȫ������״̬</option>
            <option value="1" >δ���κδ���</option>
            <option value="2" >�û��Ѿ�������</option>
            <option value="3" >�������Ѿ��յ���</option>
            <option value="4" >�������Ѿ�����</option>
            <option value="5" >�û��Ѿ��յ���</option>
          </select>
          &nbsp; 
          <input class="button" type="submit" name="Submit" value="�� ѯ">
        </div></td>
    </form>
  </tr>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
   <tr> 
	<th class="tableHeaderText" colspan=6>������Ʒ����</th>	
   </tr>
  <tr> 
    <td class="forumRowHighlight"> <div align="right"> 
        <select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" ><base target=Right> 
          <option value="editorderform.asp?zhuangtai=0" selected>--ѡ���Ѷ״̬--</option>
          <option value="editorderform.asp?zhuangtai=0" >ȫ������״̬</option>
          <option value="editorderform.asp?zhuangtai=1" >δ���κδ���</option>
          <option value="editorderform.asp?zhuangtai=2" >�û��Ѿ�������</option>
          <option value="editorderform.asp?zhuangtai=3" >�������Ѿ��յ���</option>
          <option value="editorderform.asp?zhuangtai=4" >�������Ѿ�����</option>
          <option value="editorderform.asp?zhuangtai=5" >�û��Ѿ��յ���</option>
        </select>
      </div></td>
  </tr>
</table>
  <%'��ʼ��ҳ
				Const MaxPerPage=12 
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
	if namekey="" then
	'//��״̬��ѯ
	if zhuangtai=0 or zhuangtai="" then
	select case zhuangtai
	case "0"
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 order by actiondate desc",conn,1,1
  case ""
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 order by actiondate desc",conn,1,1
  end select
  else
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where  zhuangtai="&zhuangtai&" order by actiondate",conn,1,1
  end if
  else
  '//���û���ѯ
  if zhuangtai=0 or zhuangtai="" then
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 and username='"&namekey&"' order by actiondate desc",conn,1,1
  else
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where  zhuangtai="&zhuangtai&" and username='"&namekey&"'  order by actiondate",conn,1,1
  end if
  end if
		  
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> �Բ�����ѡ���״̬Ŀǰ��û�ж�����</p>"
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
            			showpage totalput,MaxPerPage,"editorderform.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"editorderform.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"editorderform.asp"
	      				end if
	   				end if
   				   				end if

   				sub showContent
       			dim i
	   			i=0

			%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="tableBorder" >
  <tr> 
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>������</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>�µ��û�</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>����������</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>���ʽ</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>�ջ���ʽ</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>����״̬</strong></div></td>
  </tr>
  
  <%do while not rs.eof
		dim cnhww,username
		  set cnhww=server.CreateObject("adodb.recordset")
		  cnhww.open "select * from [YX_User] where id="&rs("userid"),conn,1,1
			if cnhww.eof and cnhww.bof then
	'	  response.write "�ѱ�ɾ��"
		  else
		  username=trim(cnhww("name"))
		  end if
		  cnhww.close
		  set cnhww=nothing
		  %>
  <tr bgcolor="#E8F1FF"> 
    <td height="22" class="forumRowHighlight"> <div align="center"><a href="javascript:;" onClick="javascript:window.open('vieworderform.asp?dan=<%=trim(rs("dingdan"))%>&username=<%=username%>','','width=710,height=388,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"><%=trim(rs("dingdan"))%></a></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><a href=listuser.asp?id=<%=rs("userid")%>><%=username%></a></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><%=trim(rs("userzhenshiname"))%></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"> 
        <%dim rs2
          '///֧����ʽ
          set rs2=server.CreateObject("adodb.recordset")
          rs2.open "select * from wq_songhuo where songid="&int(rs("zhifufangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "�б�ɾ���Ķ���"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.Close
          set rs2=nothing
          %>
      </div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"> 
        <%
          '///�ͻ���ʽ
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from wq_songhuo where songid="&int(rs("songhuofangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "�ѱ�ɾ��"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.close
          set rs2=nothing%>
      </div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"> 
        <%
		  select case rs("zhuangtai")
	case "1"
	response.write "δ���κδ���"
	case "2"
	response.write "�û��Ѿ�������"
	case "3"
	response.write "�������Ѿ��յ���"
	case "4"
	response.write "�������Ѿ�����"
	case "5"
	response.write "�û��Ѿ��յ���"
	end select%>
      </div></td>
  </tr>
  <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		loop
		rs.close
		set rs=nothing%>
</table>
<p> 
  <%  
				End Sub   
  
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
  				
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				'//////////////////
				if namekey="" then
				Response.Write "<form method=Post action="&filename&"?zhuangtai="&zhuangtai&">"  
				else
				Response.Write "<form method=Post action="&filename&"?zhuangtai="&zhuangtai&"&namekey="&namekey&">" 
				end if
				'//////////////////
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					'///////////////////
					if namekey="" then
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&" class='contents'>��һҳ</a> "  
					ELSE
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>��һҳ</a> "
					end if  
					'//////////////////
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else 
				'//////////////////////// 
				if namekey="" then
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&" class='contents'>βҳ</a>"
					else
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>βҳ</a>" 
					end if 
					'/////////////////////
				End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"�ʶ��� "&maxperpage&"�ʶ���/ҳ</font> " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='button' value='GO' name='cndok'></form>"  
				End Function  
			%>

</body>
</html>
