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
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<%dim zhuangtai,namekey
namekey=trim(request("namekey"))
zhuangtai=trim(request("zhuangtai"))
if zhuangtai<>"" then
if not isnumeric(zhuangtai) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
if zhuangtai="" then zhuangtai=request.QueryString("zhuangtai")
if namekey="" then namekey=request.querystring("namekey")
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
<td valign="top"> 
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
	if namekey="" then
	if zhuangtai=0 or zhuangtai="" then
	select case zhuangtai
	case "0"
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6  order by actiondate desc",conn,1,1
  case ""
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 order by actiondate desc",conn,1,1
  end select
  else
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where  zhuangtai="&zhuangtai&"  order by actiondate",conn,1,1
  end if
  else
  if zhuangtai=0 or zhuangtai="" then
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 and username='"&namekey&"'  order by actiondate desc",conn,1,1
  else
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where  zhuangtai="&zhuangtai&" and username='"&namekey&"'   order by actiondate",conn,1,1
  end if
  end if
				if err.number<>0 then
				response.write "���ݿ���������"
				end if
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> �Բ�����ѡ���״̬��û�ж�����</p>"
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
            			showpage totalput,MaxPerPage,"editorder.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"editorder.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"editorder.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0
			%>
	<table width="100%"class="tableBorder"  border="0" align="center" cellpadding="3" cellspacing="1">
      <tr>
        <td align="right" background="../images/admin_bg_1.gif"><b><font color="#ffffff">ѡ����Ҫ���ҵĶ���״̬>></font></b>
          <select name="select2" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
            <base target=Right>
            <option value="editorder.asp?zhuangtai=0" selected>--ѡ���Ѷ״̬--</option>
            <option value="editorder.asp?zhuangtai=0" >ȫ������״̬</option>
            <option value="editorder.asp?zhuangtai=1" >δ���κδ���</option>
            <option value="editorder.asp?zhuangtai=2" >�û��Ѿ�������</option>
            <option value="editorder.asp?zhuangtai=3" >�������Ѿ��յ���</option>
            <option value="editorder.asp?zhuangtai=4" >�������Ѿ�����</option>
            <option value="editorder.asp?zhuangtai=5" >�û��Ѿ��յ���</option>
          </select>          </td>
      </tr>
    </table>
	<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
	<tr > 
	<td align="center" bgcolor="fbc2c2">������</td>
	<td align="center" bgcolor="fbc2c2">�µ��û�</td>
	<td align="center" bgcolor="fbc2c2">����������</td>
	<td align="center" bgcolor="fbc2c2">���ʽ</td>
	<td align="center" bgcolor="fbc2c2">�ջ���ʽ</td>
	<td align="center" bgcolor="fbc2c2">����״̬</td>
	</tr>
        <%do while not rs.eof
		dim Godbook,username
		  set Godbook=server.CreateObject("adodb.recordset")
		  Godbook.open "select username from [user] where userid="&rs("userid"),conn,1,1
		  username=trim(Godbook("username"))
		  Godbook.close
		  set Godbook=nothing
		  %>
        <tr > 
          <td align="center"> <a href="javascript:;" onClick="javascript:window.open('vieworder.asp?dan=<%=trim(rs("dingdan"))%>&username=<%=username%>','','width=710,height=388,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"><%=trim(rs("dingdan"))%></a></td>
          <td align="center"><%=username%></td>
          <td align="center"><%=trim(rs("userzhenshiname"))%></td>
          <td align="center">
              <%dim rs2
          set rs2=server.CreateObject("adodb.recordset")
          rs2.open "select * from deliver where songid="&int(rs("zhifufangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "��ʽ�ѱ�ɾ��"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.Close
          set rs2=nothing
          %>          </td>
          <td align="center">
			<%
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from deliver where songid="&int(rs("songhuofangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "��ʽ�ѱ�ɾ��"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.close
          set rs2=nothing%>          </td>
          <td align="center">
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
	</td>        
	</tr>
        <tr >
          <td colspan="6" align="center"><%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		loop
		rs.close
		set rs=nothing%>&nbsp;</td>
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
				if namekey="" then
				Response.Write "<form method=Post action="&filename&"?zhuangtai="&zhuangtai&">"  
				else
				Response.Write "<form method=Post action="&filename&"?zhuangtai="&zhuangtai&"&namekey="&namekey&">" 
				end if
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>��ҳ ��һҳ</font> "  
				Else  
					if namekey="" then
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&" class='contents'>��һҳ</a> "  
					ELSE
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>��ҳ</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>��һҳ</a> "
					end if  
					End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>��һҳ βҳ</font>"  
				Else 
				if namekey="" then
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&" class='contents'>βҳ</a>"
					else
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>"  
					Response.Write "��һҳ</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>βҳ</a>" 
					end if 
					End If  
					Response.Write "<font class='contents'> ҳ�Σ�</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"ҳ</font> "  
					Response.Write "<font class='contents'> ����"&totalnumber&"�ʶ��� "&maxperpage&"�ʶ���/ҳ</font> " 
					Response.Write "<font class='contents'>ת����</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='contents' value='GO' name='cndok'></form>"  
				End Function  
				%>
</td>
</tr>
</table><br>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
                          <tr> 
                            <td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">������ѯ</font></b></td>
                          </tr>
                          <tr> 
                            <td height="50" > 
                              <table width="80%" border="0" align="center" cellpadding="1" cellspacing="1">
                                <tr> 
                                  <form name="form1" method="post" action="editorder.asp">
                                    <td align="center">���µ��û���ѯ 
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
                                        <input type="submit" name="Submit" value="�� ѯ">
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
