<html>
<head>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="alipayto/alipay_payto.asp"-->
<%set rss=server.CreateObject("adodb.recordset")
rss.open "select * from products where bookid="&request.querystring("id"),conn,1,3
if rss.eof or bof then
response.write "<script>alert('Sorry, not this commodity!');history.go(-1);</script>"
response.end
end if 
dim des
if not rss("metad")="" then 
des=rss("metad")
end if 
if not rss("metak")="" then 
keya=rss("metak")
end if 
%>
<%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""Illegal access!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""Illegal access!"");location.href=""index.asp"";</script>"
end if%>
<%dim bookid,action
bookid=request.QueryString("id")
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from review",conn,1,3
rs.addnew
rs("bookid")=bookid
rs("pingji")=request("pingji")
rs("pinglunname")=HTMLEncode2(trim(request("pinglunname")))
rs("pingluntitle")=HTMLEncode2(trim(request("pingluntitle")))
rs("pingluncontent")=HTMLEncode2(trim(request("pingluncontent")))
rs("ip")=Request.servervariables("REMOTE_ADDR")
rs("pinglundate")=now()
rs("shenhe")=0
rs.update
rs.close
set rs=nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from products where bookid="&bookid,conn,1,3
rs("pingji")=rs("pingji")+1
rs("pingjizong")=rs("pingjizong")+request("pingji")
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('Your comment has been successfully submitted, to be administrator audit£¡');history.go(-1);</script>"
response.End
end if
%>
<title><%=rss("bookname")%>_<%=webname%></title>
<meta name="description" content="<%=des%>">
<meta name="keywords" content="<%=keya%>"> 
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript">
	<!--
	function OpenNews() 
	{
			window.name = "news"
			win = window.open('','newswin','left=110,width=600,height=420,scrollbars=1');
	}
	//-->
	</script>
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
    <td width="190" valign="top" bgcolor="#FFFFFF">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" >
      <tr>
        <td ><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td><!--#include file="viewhistory.asp"--></td>
      </tr>
	  <tr>
        <td><!--#include file="searc.asp"--></td>
      </tr>
	  <tr>
        <td><!--#include file="shopcart.asp"--></td>
      </tr>
      <tr>
        <td><!--#include file="include/selltop.asp"--></td>
      </tr>
	  <tr>
         <td height="5" background="images/leftbg.jpg"></td>
      </tr>
	  <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width="771" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="28" background="images/body/pdbg01.gif" colspan="2"><span class="table-zuo">
          <%set rs=server.createobject("adodb.recordset")
			rs.open "select * from products where bookid="&request("id"),conn,1,3
			if rs.recordcount=0 then 
			Response.Write("<font color=""#FF0000""><strong>The merchandise has not existed</strong></font>")
			%>
          <%
			else
			bname=rs("bookname")
			rs("liulancount")=rs("liulancount")+1
			rs.update
			Sub PutIntoFav( add, ProductList )
			   If Len(ProductList) = 0 Then
				  ProductList = add
				  'ProductList = "'" & add & "'"
			   ElseIf InStr( ProductList, add ) <= 0 Then
				  ProductList = ProductList & "," & add
			   End If
			End Sub
			dim ProductList,Products
			ProductList = request.cookies("leisure")("fav")
			Products = Split(request("id"),",")
			For I=0 To UBound(Products)
			   PutIntoFav Products(I), ProductList
			Next
			response.Cookies("leisure")("fav") = ProductList
			if request.cookies("Cnhww")("username")<>"" then 
			set rs_s=server.CreateObject("adodb.recordset")
			rs_s.open "select * from [YX_User] where name='"&request.cookies("Cnhww")("username")&"'",conn,1,1
			t_userid=rs_s("id")
			rs_s.close
		    set rs_s=server.createobject("adodb.recordset")
		    rs_s.open "select * from history where bookid="&request("id")&" and username='"&request.cookies("Cnhww")("username")&"' and lx=1",conn,1,3
		    if rs_s.recordcount>0 then 
			rs_s("ltime")=now()
			rs_s("userid")=t_userid
			rs_s.update
			rs_s.close
			set rs_s=nothing
		    else
		    	rs_s.close
		    	set rs_s=server.createobject("adodb.recordset")
		    	rs_s.open "select * from history where username='"&request.cookies("Cnhww")("username")&"' and lx=1 order by ltime",conn,1,3
		    	if rs_s.recordcount>=4 then
		    	    rs_s.delete
		    	    rs_s.update
		    	end if
		    	rs_s.addnew
		    	    rs_s("username")=request.cookies("Cnhww")("username")
		    	    rs_s("bookid")=id
		    	    rs_s("bookname")=rs("bookname")
			    rs_s("userid")=t_userid
		    	    rs_s("lx")=1
		    	    rs_s("ltime")=now()
		        rs_s.update
			rs_s.close
			set rs_s=nothing
			end if
			end if
			%>
          </span>&nbsp;&nbsp;&nbsp;&nbsp;<a href=index.asp><%=webname%></a>
          <%set rs2=server.createobject("adodb.recordset")
			rs2.open "select * from bsort where anclassid="&rs("anclassid"),conn,1,1
			if rs2.recordcount>0 then
			response.write "&nbsp;&gt;&gt;&nbsp;<a href=class.asp?lx=big&anid="&rs2("anclassid")&" targer=_blank>"&rs2("anclass")&"</a> "
			end if
			rs2.close
			set rs2=server.createobject("adodb.recordset")
			rs2.open "select * from ssort where nclassid="&rs("nclassid"),conn,1,1
			if rs2.recordcount>0 then
			response.write "&nbsp;&gt;&gt;&nbsp;<a href=class.asp?lx=small&anid="&rs2("anclassid")&"&nid="&rs2("nclassid")&" targer=_blank>"&rs2("nclass")&"</a>"
			end if
			rs2.close
									
			set rs2=server.createobject("adodb.recordset")
			rs2.open "select * from xsort where id="&rs("xclassid"),conn,1,1
			if rs2.recordcount>0 then
			response.write "&gt;&gt;<a href=class.asp?lx=small&anid="&rs2("anclassid")&"&nid="&rs2("nclassid")&"&xid="&rs2("id")&" targer=_blank>"&rs2("xclass")&"</a>"
			end if 
			set rs2=nothing%></td>
      </tr>
      <tr>
        <td ><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="44%"><br>
                  <table width=300 height=300border=0 align="center" cellpadding=0 cellspacing=0>
                    <tbody>
                      <tr>
                        <td  align=center><div align="center">
                            <%if rs("zhuang")="" then 
						response.write "<img src=images/emptybook.gif width=200 border=0>"
						else%>
                            <img src="<%=trim(rs("zhuang"))%>" width=300 border=0 alt="ÉÌÆ·Ãû³Æ£º<%=rs("bookname")%>" height="300">
                            <%end if%>
                        </div></td>
                      </tr>
                    </tbody>
                </table></td>
              <td width="1%" rowspan="2" background="images/body/p_bg01.gif"></td>
              <td width="55%" rowspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="25" colspan="2" align="center">&nbsp;<font color="#ff6600" size="3"><strong><%=rs("bookname")%></strong></font></td>
                  </tr>
                  <tr>
                    <td width="28%">&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9">Item number£º</td>
                    <td width="72%" height="22"><%=trim(rs("grade"))%></td>
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif"height="7" width="9" > Specification£º</td>
                    <td height="22"><%=trim(rs("isbn"))%></td>
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif"height="7" width="9" > Brand£º</td>
                    <td height="22"><%=trim(rs("pingpai"))%></td>
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> Contrast£º</td>
                    <td height="22"><a href="to.asp?id=<%=rs("bookid")%>&action=add"><img src="images/to.gif" width="54" height="13" border="0"></a></td>
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> Number£º</td>
                    <td height="22"><%=trim(rs("kucun"))%> <%=rs("bookchuban")%>
                        <%if rs("kucun")>0 then%>
                        <font color="#999999">hot</font>
                        <%else%>
                        <font color="#999999">powered</font>
                        <%end if%>
                    </td>
                    ¡¡ 
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> Discount£º</td>
                    <td height="22"><%response.write left(rs("dazhe")*10,3)&"percent off"%>
                      (Gift of <%=rs("yeshu")%>) </td>
                  </tr>
                  <tr>
                    <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> views£º</td>
                    <td height="22"><%=trim(rs("liulancount"))%></td>
                  </tr>
                  <form method=post action=buy.asp?id=<%=rs("bookid")%>&action=add>
                    <tr>
                      <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> Size£º</td>
                      <td height="22"><%if rs("chima")<>"" then%>
                          <%size=trim(rs("chima"))%>
                          <%myarr=split(size,"/")%>
                          <select name="chima" id="chima">
                            <% for i=0 to ubound(myarr) %>
                            <option value="<%=trim(myarr(i))%>" selected><%=trim(myarr(i))%></option>
                            <%next%>
                          </select>
                          <%else
						response.write "No Size information" %>
                          <%end if%></td>
                    </tr>
                    <tr>
                      <td>&nbsp;<img src="images/body/orange-bullet.gif" height="7" width="9"> Color£º</td>
                      <td height="22"><%if rs("yanse")<>"" then%>
                          <%size=trim(rs("yanse"))%>
                          <%myarr=split(size,"/")%>
                          <select name="yanse" id="select">
                            <% for i=0 to ubound(myarr) %>
                            <option value="<%=trim(myarr(i))%>" selected><%=trim(myarr(i))%></option>
                            <%next%>
                          </select>
                          <%else
					  response.write "No color information" %>
                          <%end if%></td>
                    </tr>
                    <tr>
                      <td height="8" colspan="2" background="images/body/p_bg02.gif"></td>
                    </tr>
                    <tr>
                      <td width="24%" height="13" align="right"><div align="right">market price£º</div></td>
                      <td width="76%"><s>$<%=rs("shichangjia")%></s></td>
                    </tr>
                    <tr align="right">
                      <td height="2" colspan="2" background="images/body/p_bg03.gif"></td>
                    </tr>
                    <tr>
                      <td align="right" ><div align="right">Member price£º</div></td>
                      <td ><font color="#FF0000">$<%=trim(rs("huiyuanjia"))%> </font></td>
                    </tr>
                    <tr>
                      <td align="right" ><div align="right">VIP price£º</div></td>
                      <td ><font color="#FF0000">
                        <%if Request.Cookies("cnhww")("reglx")=2 then%>
                        $<%=trim(rs("Vipjia"))%>
                        <%else%>
                        <font color="#FF3333">More concessions...</font>
                        <%end if%>
                      </font>&nbsp;</td>
                    </tr>
                    <tr>
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td  height="10"></td>
                          </tr>
                          <tr>
                            <td><input name="image" type="image" src="images/gouwuche.jpg" width="185" height="36"  border=0>
                                <% if  alipaytype=0 then %>
                              &nbsp;<a href="alipay.asp?p=<%  if Request.Cookies("cnhww")("reglx")=2 then response.write rs("vipjia") else response.write rs("huiyuanjia") end if %>&d=<%=rs("bookname")%>" target="_blank"><img src=images/zhifubao.jpg border=0 width=98 height=36></a>
                                <%end if %>
                              &nbsp;<a href="javascript:;" onClick="javascript:window.open('stow.asp?id=<%=rs("bookid")%>&action=add','','width=640,height=260,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"><img src="images/shoucang.jpg" width="160" height="36"  border=0></a> </td>
                          </tr>
                      </table></td>
                    </tr>
                  </form>
              </table></td>
            </tr>
            <tr>
              <td align="center"><a href="javascript:;" onClick="javascript:window.open('<%=trim(rs("zhuang"))%>','','width=400,height=400,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"> <img src="images/body/itemzoom.gif" width="114" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td rowspan="2">&nbsp;</td>
      </tr>
    </table>
      <table  style=" BORDER-bottom:#E5E5E5 1px solid; BORDER-top:#E5E5E5 1px solid;BORDER-left:#E5E5E5 1px solid; BORDER-right:#E5E5E5 1px solid;" width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		
              <tr>
              <td height="17" colspan="2" ><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="left"><img src="images/body/p_dtitle01.gif" width="770" height="29" ></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td colspan="2"><table width="98%" border="0" cellspacing="0" cellpadding="10" align="center">
                  <tr>
                    <td>
                        <%=rs("bookcontent")%> </td>
                  </tr>
				  <tr>
				<td align=right> >>See more about <%
					response.write "<a href=research.asp?action=1&anclassid=0&searchkey="&trim(rs("bookname"))&" ><font color=red>"&trim(rs("bookname"))&"</font></a> " end if %>Product&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</td>
				</tr>
              </table>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                    <td><img src="images/tltj.gif" width="770"  height="29"></td>
                  </tr><tr> 
                      <td valign="middle"><table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
                          <%
						  sql="select top 5 * from products where anclassid = (select anclassid from products where bookid = "&request("id")&") order by liulancount ,chengjiaocount,adddate desc " 
						  set rs=server.createobject("adodb.recordset")
						  rs.open sql,conn,1,1
						%><tr> 
                            <%
							  if rs.eof and rs.bof then
							  response.write "<td  align=center>Sorry, this column is not suitable for the goods you buy£¡</td>"
							  'response.End
							  else
							  %>
                            <%
							if not rs.eof then
							i=1
							do while not rs.eof%>
                            <td width="195" height="135" align="center" valign="top">
                              <table width="20%"  border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td><div align="center"><a href=products.asp?id=<%=rs("bookid")%> target="_blank" class="style1"> <%=InterceptString(rs("bookname"),16)%></a></div></td>
                                </tr>
                                <tr> 
                                  <td><TABLE width=90 height=90 border=0 align="center" cellPadding=2 cellSpacing=1>
                                      <TBODY>
                                        <TR> 
                                          <TD align=center> 
                                            <%if rs("bookpic")="" then 
											response.write "<div align=center><a href=products.asp?id="&rs("bookid")&" > <img src=images/emptybook.gif width=90 height=90 border=0></a></div>"
											else%>
                                            <a href=products.asp?id=<%=rs("bookid")%> target="_blank"><img src="<%=trim(rs("bookpic"))%>" alt="<%=rs("bookname")%>" width=95 height=95 border=0 align=absmiddle></a> 
                                            <%end if%>                                          </TD>
                                        </TR>
                                      </TBODY>
                                    </TABLE></td>
                                </tr>
                              </table></td>
                            <%if i mod 5 = 0 then%>
                          </tr>
                          <%end if
	    rs.movenext
         i=i+1
    	 loop
		rs.close
		end if
		end if 
	 
	%>
                        </table></td>
                    </tr>
                </table> </td>
            </tr>
            <tr>
              <td colspan="2" ><img src="images/body/views.gif" width="770"  height="29"></td>
            </tr>
			<tr>
              <td colspan="2" align=right>&nbsp;</td>
            </tr>
            <%
			set rs=server.createobject("adodb.recordset")
			rs.open "select  * from review where bookid="&request("id")&"and shenhe=1 order by pinglundate desc",conn,1,1
			j=rs.recordcount
			plcount=rs.recordcount
			if j>3 then j=3
			for i=1 to j
			%>
            <tr>
              <td colspan="2" valign="middle">
                <table width="95%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#e1e1e1">
                <tr bgcolor="#FFFFFF">
                  <td height="15" colspan="2">Title£º <%=trim(rs("pingluntitle"))%></td>
                  <td width="26%" height="15"><img src="images/level<%=rs("pingji")%>.gif"></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="15" colspan="2">Author£º <%=trim(rs("pinglunname"))%>¡¡(<%=trim(rs("pinglundate"))%>)</td>
                  <td width="26%" height="15">Views Ip£º<%=trim(rs("ip"))%></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td height="15" colspan="3" bgcolor="#FFFFFF">Content£º
                    <%if len(rs("pingluncontent"))>24 then 
	response.write left(rs("pingluncontent"),20)&"... <a href=""javascript:;"" onClick=""javascript:window.open('pinglunll.asp?id="&rs("pinglunid")&"','pinglun','width=300,height=400,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');return false;"" title="&trim(rs("pingluntitle"))&">[More]</a>" 
	else response.write rs("pingluncontent")
	end if%>                  </td>
                </tr>
                <%if trim(rs("huifu"))<>"" then%>
                <tr bgcolor="#FFFFFF">
                  <td height="15" colspan="3"><b><font color="#ff6600">Sellers reply£º</font></b><%=rs("huifu")%> </td>
                </tr>
                <%end if%>
              </table>             </td>
            </tr>
            <%	
rs.movenext
next
if plcount>0 then 
%>
            <tr>
              <td colspan="2" align=right> >>Browse products <a href=reviewlist.asp?id=<%=request("id")%> > All <font color="#cc0000"><%=plcount%></font> Comment</a> &nbsp;&nbsp;&nbsp;&nbsp;</td>
            </tr>
            <%end if%>
            <tr>
              <td height="180" colspan="2" align="center"><!--#include file="review.asp"--></td>
            </tr>
      </table><br></td>
    </tr>
  </table>
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
</table>
</div>
<!--#include file="include/foot.asp"-->
</body>
</html>
  <%
function InterceptString(txt,length) 
txt=trim(txt) 
x = len(txt) 
y = 0 
if x >= 1 then 
for ii = 1 to x 
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then 

y = y + 2 
else 
y = y + 1 
end if 
if y >= length then 
txt = left(trim(txt),ii) 

exit for 
end if 
next 
InterceptString = txt 
else 
InterceptString = "" 
end if 

End Function 
%>