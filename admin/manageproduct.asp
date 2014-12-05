<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
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

<%
dim selectm,selectkey,selectbookid
selectkey=trim(request(trim("selectkey")))
selectm=trim(request("selectm"))
if selectkey="" then
selectkey=request.QueryString("selectkey")
end if
if selectkey="请输入关键字" then
selectkey=""
end if
if selectm="" then
selectm=request.QueryString("selectm")
end if
selectbookid=request("selectbookid")
if selectbookid<>"" then
conn.execute "delete from products where bookid in ("&selectbookid&")"
response.Redirect "manageproduct.asp"
response.End
end if
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
bid=trim(request("bid"))
			if bid="" then
           		rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic from products order by adddate desc",conn,1,1
				elseif bid=0 then
			  rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic from products order by adddate desc",conn,1,1
              else
		   		rs.open "select * from products where anclassid="&bid&" order by adddate desc",conn,1,1
		
				end if 
		    case "0"
			rs.open "select bookid,bookname,adddate,zhuangtai,kucun,bookpic  from products order by adddate desc",conn,1,1
		    case "bookid"
			if selectkey="" then
			selectkey=0
			end if
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where bookid="&selectkey&" order by adddate desc",conn,1,1
			case "bookname"
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where bookname like '%"&selectkey&"%' order by adddate desc",conn,1,1
			case "bookcontent"
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where bookcontent like '%"&selectkey&"%' order by adddate desc",conn,1,1
			case "news"
			if selectkey<>"" then
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where newsbook=1 and bookname like '%"&selectkey&"%' order by adddate desc",conn,1,1
			else
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where newsbook=1 order by adddate desc",conn,1,1
			end if
			case "tuijian"
			if selectkey<>"" then
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where bestbook=1 and bookname like '%"&selectkey&"%' order by adddate desc",conn,1,1
			else
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where bestbook=1 order by adddate desc",conn,1,1
			end if
			case "tejia"
			if selectkey<>"" then
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic  from products where tejiabook=1 and bookname like '%"&selectkey&"%' order by adddate desc",conn,1,1
			else
			rs.open "select bookid,bookname,adddate,kucun,zhuangtai,bookpic from products where tejiabook=1 order by adddate desc",conn,1,1
			end if
		  	end select
		   	if err.number<>0 then
				response.write "数据库中无数据"
				end if
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> 数据库中无数据！</p>"
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
            			showpage totalput,MaxPerPage,"manageproduct.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"manageproduct.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"manageproduct.asp"
	      				end if
	   				end if
   				   	end if
   				sub showContent
       				dim i
	   			i=0%>
        <table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
          <tr> 
            <td colspan="7" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商品查看与修改
              <select name="zhuangtai" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
                <base target=Right>
                <option value="" selected>--请选择商品分类--</option>
                <option value="manageproduct.asp?bid=0" >--显示所有商品--</option>
                <% set vs=server.createobject("adodb.recordset")
		  vs.open "select * from bsort ",conn,1,1 
		  do while not vs.eof  %>
                <option value="manageproduct.asp?bid=<%=trim(vs("anclassid"))%>" ><%=vs("anclass")%></option>
                <%vs.movenext
	  loop
	  %>
              </select>
            </font></b></td>
          </tr>
          <tr > 
            <td width="12%" align="center" bgcolor="fbc2c2">序号</td>
            <td width="12%" align="center" bgcolor="fbc2c2">图片</td>
            <td width="26%" align="center" bgcolor="fbc2c2">商品名称</td>
            <td width="12%" align="center" bgcolor="fbc2c2">商品状态</td>
            <td width="9%" align="center" bgcolor="fbc2c2">复制商品</td>
            <td width="18%" align="center" bgcolor="fbc2c2">加入时间</td>
            <td width="11%" align="center" bgcolor="fbc2c2">选 择</td>
          </tr>
          <%do while not rs.eof%>
          <tr > 
            <td align="center"><%=rs("bookid")%></td>
            <td align="center"> 
              <%if rs("bookpic")="" then 
	response.write "<div align=center><a href=../products.asp?id="&rs("bookid")&" target=_blank><img src=../images/emptybook.gif width=60 height=60 border=0></a></div>"
	else%>
              <a href=../products.asp?id=<%=rs("bookid")%> target=_blank><img src="../<%=trim(rs("bookpic"))%>" width=60 height=60 border=0 align=absmiddle></a> 
              <%end if%>            </td>
            <td width="26%" STYLE='PADDING-LEFT: 10px'> <a href=editproduct.asp?id=<%=rs("bookid")%>> 
              <% if len(trim(rs("bookname")))>20 then
			response.write left(trim(rs("bookname")),18)&".."
			else
			response.write trim(rs("bookname"))
			end if%>
              </a> 
              <%if rs("kucun")>0 then%>
              <font color=blue> 热卖中</font> 
              <%else%>
              <font color=red> 缺货</font> 
              <%end if%>            </td>
            <td width="12%" align="center" STYLE='PADDING-LEFT: 10px'> 
              <% if rs("zhuangtai")=0 then %>
              <a href="keep.asp?action=xj&id=<%=rs("bookid")%>" title="点击可使此商品下架">上架中</a> 
              <%else  %>
              <a href="keep.asp?action=sj&id=<%=rs("bookid")%>" title="点击可使此商品上架"><font color="#FF0000">下架中</font></a> 
              <%end if%>            </td>
            <td align="center"><a href="copypro.asp?id=<%=rs("bookid")%>" onClick="return confirm('确定复制 <%=rs("bookname")%> 吗？')"><img src="images/icon_copy.gif" alt="复制此商品，对于参数相同的商品可以进行快速修改和编辑。" border=0></a></td>
            <td align="center"> <%=rs("adddate")%></td>
            <td align="center"> <input name="selectbookid" type="checkbox" id="selectbookid" value="<%=rs("bookid")%>">            </td>
          </tr>
          <tr > 
            <td height="1" align="center"></td>
            <td height="1" colspan="5" align="center" bgcolor="#CCCCCC"></td>
            <td height="1" align="center"></td>
          </tr>
          <%i=i+1
		if i>=MaxPerPage then Exit Do
		rs.movenext
		loop
		rs.close
		set rs=nothing%>
          <tr > 
            <td height="30" colspan="7" align="right">全选 
              <input type="checkbox" name="checkbox" value="Check All" onClick="mm()"> 
              <input type="submit" name="Submit" value="删 除" onClick="return test();"> 
              &nbsp;</td>
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
					Response.Write "<font class='contents'>首页 上一页</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>首页</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>上一页</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>下一页 尾页</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>"  
					Response.Write "下一页</a> <a href="&filename&"?page="&n&"&selectm="&selectm&"&selectkey="&selectkey&" class='contents'>尾页</a>"  
				End If  
					Response.Write "<font class='contents'> 页次：</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"页</font> "  
					Response.Write "<font class='contents'> 共有"&totalnumber&"种商品 " 
					Response.Write "<font class='contents'>转到：</font><input type='text' name='page' size=2 maxlength=10 class=smallInput value="&currentpage&">"  
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
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商品查讯</font></b></td>
</tr>
<tr > 
<form name="form2" method="post" action="manageproduct.asp">
<td height="50"> 
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td width="35%" align="center">
<input name="selectkey" type="text" id="selectkey" onFocus="this.value=''" value="请输入关键字">
</td>
<td width="45%" align="center"> 
<select name="selectm" id="selectm">
<option value="bookname">按商品名称</option>
<option value="bookcontent">按商品说明</option>
<option value="bookid">按商品序号</option>
<option value="0">全部商品</option>
<option value="news">按新品</option>
<option value="tejia">按特价</option>
<option value="tuijian">按推荐</option>
</select>
</td>
<td width="20%" height="30" align="center">
<input type="submit" name="Submit2" value="查 讯">
</td>
</tr>
<tr><td colspan=3 height=40>　除查询“所有商品”外，必须要输入关键字，才能查询<br>　根据商品序号来查询时，关键字里只能输入数字。</td></tr>
</table>
</td>
</form>
</tr>
</table>
<br>
<!--#include file="foot.asp"-->
</body>
</html>
<script>
function test()
{
  if(!confirm('确认删除吗？')) return false;
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
