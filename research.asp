<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%if IsNumeric(request.QueryString("jiage"))=False then
        response.write"<script>alert(""Input error, if in doubt please contact the administrator£¡"");location.href=""index.asp"";</script>"
        response.end
	end if
dim action,searchkey,anclassid,jiage,selectname
anclassid=request("anclassid")
searchkey=request("searchkey")
jiage=request("jiage")
action=request("action")
selectname=request("selectname")
if searchkey<>"" then
set rw=server.CreateObject("adodb.recordset")
			rw.open "select * from words where keyword='"&searchkey&"'",conn,1,3
			if rw.eof then
			rw.addnew
			rw("keyword")=searchkey
			rw("ip")=Request.ServerVariables("REMOTE_ADDR")
			rw("times")=1
			rw.update
			else
			rw("times")=rw("times")+1
			rw.update
			rw.close
			set rw=nothing
			end if 
			end if 
if InStr(searchkey,"'")>0 then
response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
response.end
end if
if anclassid<>"" then
if not isnumeric(anclassid) then 
response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
if jiage<>"" then
if not isnumeric(jiage) then 
response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
if action<>"" then
if not isnumeric(action) then 
response.write"<script>alert(""Illegal access!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
%>
<html><head><title><%=webname%>--Product search results</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td background="images/leftbg.jpg"><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg"></td>
      </tr>
      
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td background="images/body/pdbg01.gif" height=28> 
            <%if jiage="" then%>
            <%if searchkey="" then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;All goods"
	      else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;Regarding the goods you are looking for£º¡¡"
		if anclassid<>0 then
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select * from bsort where anclassid="&anclassid,conn,1,1
			response.write "<a href=class.asp?lx=big&anid="&anclassid&"><font color=red>"&rs("anclass")&"</font></a>"&" &gt;&gt; "
			rs.close
			set rs=nothing
		end if
		response.write "<font color=red>"&searchkey&"</font>"
		s_bookname=searchkey
	      end if%>
            <%else%>
            <%if (action="1" or action="3") and searchkey="" then
		response.write "<script language=javascript>alert('Sorry, please enter your inquiry keyword');history.go(-1);</script>"
		response.End
	      else
           if action="2" then
           set rs=server.createobject("adodb.recordset")
           rs.open "select * from brand where tuijian=1 order by pingpaiorder",conn,1,1
           if rs.recordcount>0 then
           
           for i=1 to rs.recordcount
         
           rs.movenext
           next
           rs.close
           set rs=nothing
         
           end if
           end if
	        response.write "&nbsp;&nbsp;&nbsp;&nbsp;Regarding the goods you are looking for£º¡¡"
		    if anclassid<>0 then
			set rs=server.CreateObject("adodb.recordset")
			rs.open "select * from bsort where anclassid="&anclassid,conn,1,1
			response.write "<a href=class.asp?lx=big&anid="&anclassid&"><font color=red>"&rs("anclass")&"</font></a>"&" &gt;&gt; "
			rs.close
			set rs=nothing
		    end if
		if action="1" or action="3" then 
			response.write "<font color=red>"&searchkey&"</font>"
			s_bookname=searchkey
		else
			response.write "<font color=red>"&selectname&"</font>"
			s_bookname=selectname
end if
end if%>
            <%end if%></td>
      </tr>
	  <tr><td height="15"></td></tr>
    </table>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
      <tr>
        <td align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><%
            	if request("search")=" all" then 
				keyword=request.form("keyword")
				if keyword<>"" then
				dim path,myfileobject,mytextfile
				 path = Server.MapPath(shopurl)
				set  MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
				set  MyTextFile = MyFileObject.CreateTextFile(path)
				MyTextFile.WriteLine (keyword)
				MyTextFile.Close()
				response.end
					else
				response.end
					end if 
					else
					end if 
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
				if jiage="" then  
					sql2="searchkey="&searchkey&"&anclassid="&anclassid
						select case request("anclassid")
						case "0"
						sql1=" bookname like '%"&searchkey&"%' "
						case else
						sql1=" bookname like '%"&searchkey&"%' and anclassid="&request("anclassid")&" "
						end select
				else
					sql2="searchkey="&searchkey&"&anclassid="&anclassid&"&jiage="&jiage&"&action="&action&"&selectname="&selectname
						if anclassid<>0 then   
						select case action
						case "1"
						sql1=" bookname like '%"&searchkey&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") and anclassid="&anclassid&" "
						case "2"
						sql1=" pingpai like '%"&selectname&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") and anclassid="&anclassid&" "
						case "3"
						sql1=" bookcontent like '%"&searchkey&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") and anclassid="&anclassid&" "
						end select
						else
						select case action
						case "1"
						sql1=" bookname like '%"&searchkey&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") "
						case "2"
						sql1=" pingpai like '%"&selectname&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") "
						case "3"
						sql1=" bookcontent like '%"&searchkey&"%' and (shichangjia<"&jiage&" or huiyuanjia<"&jiage&" or vipjia<"&jiage&") "
						end select
						end if
				end if
				call sss()
				set rs=server.CreateObject("adodb.recordset")
				rs.open "select * from products where "&sql1&" and zhuangtai=0 order by adddate desc",conn,1,1	
  				if rs.eof And rs.bof then
       				Response.Write "<p align=center> Sorry, no inquiry to you need goods£¡</p><br>"
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
            			showpage totalput,MaxPerPage,"research.asp"
       				else
          				if (currentPage-1)*MaxPerPage<totalPut then
            				rs.move  (currentPage-1)*MaxPerPage
            				dim bookmark
            				bookmark=rs.bookmark
            				showContent
             				showpage totalput,MaxPerPage,"research.asp"
        				else
	        				currentPage=1
           					showContent
           					showpage totalput,MaxPerPage,"research.asp"
	      				end if
	   				end if
   				   				end if
   				sub showContent
       			dim i
	   			i=0
			%>
          <%do while not rs.eof%>
           
            <table width="95%" border="0" cellspacing="0" cellpadding="0"  >
              <tr>
                <td width="22%"><table align=center cellspacing=0 cellpadding=0 width=100 height=100 border=0>
                  <tbody>
                    <tr>
                        <td  align=center> 
                          <%if rs("bookpic")="" then 
response.write "<div align=center><a href=products.asp?id="&rs("bookid")&" ><img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
else%>
                          <a href=products.asp?id=<%=rs("bookid")%> ><img src="<%=trim(rs("bookpic"))%>"  width=105 border=0 height="105"></a> 
                          <%end if%>                        </td>
                    </tr>
                  </tbody>
                </table></td>
                <td width="35%" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>Name£º
                      <%response.write "<a href=products.asp?id="&trim(rs("bookid"))&" >"&trim(rs("bookname"))&"</a>"%></td>
                    </tr>
                    
                  </table></td>
                <td width="22%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="25">market price $£º<s><%=rs("shichangjia")&""%></s> </td>
                    </tr>
                    <tr>
                      <td height="25">Member price $£º<%="<font color=#FF0000>"&rs("huiyuanjia")&"</font>"%></td>
                    </tr>
                  </table></td>
                <td width="21%" align="center"><a href="buy.asp?id=<%=rs("bookid")%>&action=add" target="_blank" ><img src=images/buy.gif width=91 height=30 border=0></a><br>
                <a href=# onClick="window.open('stow.asp?id=<%=rs("bookid")%>&action=add ', 'basket','menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbars=0,resizable=0,width=520,top=10,left=0,height=260')"><img src='images/house.gif' width=91 height=30 border=0></a></td>
              </tr>
			  <tr>
                <td height="8" colspan="4"></td>
              </tr>
              <tr>
                <td height="1" colspan="4" background="images/mingle/inbg.gif"></td>
              </tr>
			  <tr>
                <td height="8" colspan="4"></td>
              </tr>
            </table>
            <%i=i+1
		if i>=MaxPerPage then Exit Do
		rs.movenext
  loop
  rs.close
  set rs=nothing%>
            <%  
				End Sub   
				Function showpage(totalnumber,maxperpage,filename)  
  				Dim n
				If totalnumber Mod maxperpage=0 Then  
					n= totalnumber \ maxperpage  
				Else
					n= totalnumber \ maxperpage+1  
				End If
				Response.Write "<form method=Post action="&filename&"?searchkey="&searchkey&"&action="&action&"&anclassid="&anclassid&">"  
				Response.Write "<p align='center' class='contents'> "  
				If CurrentPage<2 Then  
					Response.Write "<font class='contents'>Home&nbsp;&nbsp;Previous</font> "  
				Else  
					Response.Write "<a href="&filename&"?page=1&searchkey="&searchkey&"&action="&action&"&anclassid="&anclassid&" class='contents'>Home</a> "  
					Response.Write "<a href="&filename&"?page="&CurrentPage-1&"&searchkey="&searchkey&"&action="&action&"&anclassid="&anclassid&" class='contents'>Previous</a> "  
				End If
				If n-currentpage<1 Then  
					Response.Write "<font class='contents'>Next&nbsp;&nbsp;end</font>"  
				Else  
					Response.Write "<a href="&filename&"?page="&(CurrentPage+1)&"&searchkey="&searchkey&"&action="&action&"&anclassid="&anclassid&" class='contents'>"  
					Response.Write "Next</a>&nbsp;&nbsp;<a href="&filename&"?page="&n&"&searchkey="&searchkey&"&action="&action&"&anclassid="&anclassid&"&jiage="&jiage&" class='contents'>End</a>"  
				End If  
					Response.Write "<font class='contents'> Ones£º</font><font class='contents'>"&CurrentPage&"</font><font class='contents'>/"&n&"Page</font> "  
					Response.Write "<font class='contents'> All Found"&totalnumber&"Goods " 
					Response.Write "<font class='contents'>Go to£º</font><input CLASS='wenbenkuang' type='text' name='page' size=2 maxlength=8 class=smallInput value="&currentpage&">"  
					Response.Write "&nbsp;<input type='submit' CLASS='go-wenbenkuang' class='contents' value='GO' name='cndok'></form>"  
				End Function  
			%>
          <%
sub sss() 
if request.cookies("Cnhww")("username")<>"" then 
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from [user] where username='"&request.cookies("Cnhww")("username")&"'",conn,1,1
t_userid=rs_s("userid")
rs_s.close
set rs_s=server.CreateObject("adodb.recordset")
rs_s.open "select * from history where username='"&request.cookies("Cnhww")("username")&"' and bookid='"&s_bookid&"' bookname='"&s_bookname&"' and lx=2",conn,1,3
if rs_s.recordcount>0 then 
	rs_s("ltime")=now()
	rs_s("userid")=t_userid
	rs_s("searchkey")=sql2
	rs_s.update
	rs_s.close
	set rs_s=nothing
else
		    	rs_s.close
		    	set rs_s=server.createobject("adodb.recordset")
		    	rs_s.open "select * from history where username='"&request.cookies("Cnhww")("username")&"' and lx=2",conn,1,3
		    	if rs_s.recordcount>=4 then
		    	    rs_s.delete
		    	    rs_s.update
		    	end if
		    	rs_s.addnew
		    	    rs_s("username")=request.cookies("Cnhww")("username")
		    	    rs_s("searchkey")=sql2
		    	    rs_s("bookname")=s_bookname
					rs_s("bookid")=s_bookid
			    rs_s("userid")=t_userid
		    	    rs_s("lx")=2
		    	    rs_s("ltime")=now()
		        rs_s.update
			rs_s.close
			set rs_s=nothing
		    end if
		end if
end sub%>        </td>
      </tr>
    </table></td>
  </tr>
  <tr><td height="40" colspan="2">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
