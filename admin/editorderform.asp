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
	<th class="tableHeaderText" colspan=6>定单查询</th>	
   </tr>
  <tr> 
    <form name="form1" method="post" action="editorderform.asp">
      <td class="forumRowHighlight"> <div align="center">按下单用户查询 
          <input name="namekey" type="text" id="namekey" value="请输入用户名" size="14" onFocus="this.value=''">
          &nbsp; 
          <select name="zhuangtai" id="zhuangtai">
            <option value="0" selected>--选择查询状态--</option>
            <option value="0" >全部订单状态</option>
            <option value="1" >未作任何处理</option>
            <option value="2" >用户已经划出款</option>
            <option value="3" >服务商已经收到款</option>
            <option value="4" >服务商已经发货</option>
            <option value="5" >用户已经收到货</option>
          </select>
          &nbsp; 
          <input class="button" type="submit" name="Submit" value="查 询">
        </div></td>
    </form>
  </tr>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
   <tr> 
	<th class="tableHeaderText" colspan=6>管理商品订单</th>	
   </tr>
  <tr> 
    <td class="forumRowHighlight"> <div align="right"> 
        <select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" ><base target=Right> 
          <option value="editorderform.asp?zhuangtai=0" selected>--选择查讯状态--</option>
          <option value="editorderform.asp?zhuangtai=0" >全部订单状态</option>
          <option value="editorderform.asp?zhuangtai=1" >未作任何处理</option>
          <option value="editorderform.asp?zhuangtai=2" >用户已经划出款</option>
          <option value="editorderform.asp?zhuangtai=3" >服务商已经收到款</option>
          <option value="editorderform.asp?zhuangtai=4" >服务商已经发货</option>
          <option value="editorderform.asp?zhuangtai=5" >用户已经收到货</option>
        </select>
      </div></td>
  </tr>
</table>
  <%'开始分页
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
	'//按状态查询
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
  '//按用户查询
  if zhuangtai=0 or zhuangtai="" then
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where zhuangtai<6 and username='"&namekey&"' order by actiondate desc",conn,1,1
  else
  rs.open "select distinct(dingdan),userid,userzhenshiname,actiondate,songhuofangshi,zhifufangshi,zhuangtai from orders where  zhuangtai="&zhuangtai&" and username='"&namekey&"'  order by actiondate",conn,1,1
  end if
  end if
		  
				if err.number<>0 then
				response.write "数据库中无数据"
				end if
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 对不起，您选择的状态目前还没有订单！</p>"
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
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>订单号</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>下单用户</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>订货人姓名</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>付款方式</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>收货方式</strong></div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"><strong>订单状态</strong></div></td>
  </tr>
  
  <%do while not rs.eof
		dim cnhww,username
		  set cnhww=server.CreateObject("adodb.recordset")
		  cnhww.open "select * from [YX_User] where id="&rs("userid"),conn,1,1
			if cnhww.eof and cnhww.bof then
	'	  response.write "已被删除"
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
          '///支付方式
          set rs2=server.CreateObject("adodb.recordset")
          rs2.open "select * from wq_songhuo where songid="&int(rs("zhifufangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "有被删除的订单"
		  else
          response.Write trim(rs2("subject"))
          end if
		  rs2.Close
          set rs2=nothing
          %>
      </div></td>
    <td height="22" class="forumRowHighlight"> <div align="center"> 
        <%
          '///送货方式
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select * from wq_songhuo where songid="&int(rs("songhuofangshi")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "已被删除"
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
	response.write "未作任何处理"
	case "2"
	response.write "用户已经划出款"
	case "3"
	response.write "服务商已经收到款"
	case "4"
	response.write "服务商已经发货"
	case "5"
	response.write "用户已经收到货"
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
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					'///////////////////
					if namekey="" then
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&" class='contents'>上一页</a> "  
					ELSE
					Response.Write "<a href="&filename&"?page=1&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>上一页</a> "
					end if  
					'//////////////////
				End If
				
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else 
				'//////////////////////// 
				if namekey="" then
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&" class='contents'>尾页</a>"
					else
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&zhuangtai="&zhuangtai&"&namekey="&namekey&" class='contents'>尾页</a>" 
					end if 
					'/////////////////////
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"笔订单 "&maxperpage&"笔订单/页</font> " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit'  class='button' value='GO' name='cndok'></form>"  
				End Function  
			%>

</body>
</html>
