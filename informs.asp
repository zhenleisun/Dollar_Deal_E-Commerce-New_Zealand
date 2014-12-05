<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html><head><title><%=webname%>--行业资讯</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<meta name="keywords" content="网趣网上购物系统,网趣网上购物系统时尚版,网趣购物系统,网上购物系统,购物系统,网趣购物,商城源码,网上商店,网上商店系统,域名注册,虚拟主机,恒伟网络">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%if IsNumeric(request.QueryString("id"))=False then
response.write("<script>alert(""非法访问!"");location.href=""index.asp"";</script>")
response.end
end if
dim id
id=request.QueryString("id")
if not isinteger(id) then
response.write"<script>alert(""非法访问!"");location.href=""index.asp"";</script>"
end if%>
     <%dim newsid
newsid=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from inform where newsid="&newsid,conn,1,3
rs("viewcountent")=rs("viewcountent")+1
rs.update
%>
<!--#include file="include/head.asp"-->
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="39" valign="top"><table width="100%" height="184" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="right" valign="top"><img src="images/body/jiao.gif" width="15" height="17"></td>
      </tr>
      <tr>
        <td height="90"><div align="right"><img src="images/ttts.gif" width="26" height="110" /></div></td>
      </tr>
      <tr>
        <td height="30">&nbsp;</td>
      </tr>
    </table></td>
    <td width="190"style="BORDER-bottom:#FF67A0 1px solid;BORDER-left:#FF67A0 1px solid; BORDER-right:#FF67A0 1px solid;" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
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
        <td><!--#include file="searc.asp"--><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td></td>
      </tr>
    </table></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
          <td background="images/body/pdbg01.gif" height=28>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <a href=index.asp><%=webname%></a> >> <a href="inform.asp">行业资讯</a></td>
      </tr>
	  <tr>
        <td ><%set rs=server.createobject("adodb.recordset")
		rs.open "select * from inform where newsid="&request("id"),conn,1,1
		title=trim(rs("title"))
		UpdateTime=trim(rs("UpdateTime"))
		News_Content=rs("Content")
		view=rs("viewcountent")
			rs.Close
		set rs=nothing
			%></td>
      </tr>
      <tr>
        <td><br>
            <table width="92%" height="26"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center"><font color="#F46404"> 
                  <%
tmp_center="<strong><font face="" color="">"
if titletype="0" then
tmp_center=tmp_center&title
end if
if titletype="1" then
tmp_center=tmp_center&"<b><i>"&title&"</i></b>"
end if
if titletype="b" then
tmp_center=tmp_center&"<b>"&title&"</b>"
end if
if titletype="i" then
tmp_center=tmp_center&"<i>"&title&"</i>"
end if
if titletype="u" then
tmp_center=tmp_center&"<u>"&title&"</u>"
end if
tmp_center=tmp_center&"</font></strong>"&VbCrLf
response.write tmp_center
%>
                  <br>
                  <br>
                  <br>
                  浏览 <font color=red><%=view%></font> 
                  次【字号 <A     onclick="Zoom.style.fontSize='19px';" href="#">大</A> 
                  <A 
            onclick="Zoom.style.fontSize='16px';" 
            href="#">中</A> <A    onclick="Zoom.style.fontSize='12px';" 
            href="#">小</A>】 发布时间：<%=updatetime%>
                  <a class="b12" href="#"
            onclick="if (window.print) window.print();return false">打印本页</a> </font>
                  <hr width="500" size="1" noshade>
                  <font color="#F46404">&nbsp; </font> </td>
              </tr>
            </table>
          <table width="90%" border="0" cellspacing="0" cellpadding="5" align="center">
              <tr>
                <td><DIV id=Content><FONT id=Zoom><%=HtmlSelfEnCode(News_Content)%></div></td>
              </tr>
            </table>
          <br>
          <table width="580" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height=1 colspan="3"background="images/mingle/inbg.gif"></td>
            </tr>
            <tr>
              <td height="40" align="right">
			  <a href='javascript:onclick=history.go(-1)'><img src="images/mingle/infbk.gif" width="108" height="24" border="0" ></a></td>
            </tr>
          </table>
          <%
Function HtmlSelfEnCode(content)
TempContent=content
TempContent=replace(TempContent,"<img src=","<P align=center><img  onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333' src=")
HtmlSelfEnCode=TempContent
End Function
%></td>
      </tr>
    </table></td>
    <td width="68" style="BORDER-left:#FF67A0 1px solid;">&nbsp;</td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>