<!--#include file="conn.asp"-->
<%
if request("delid")<>"" then call delmsg()
%>

<HTML>
<HEAD>
<TITLE>站内消息-<%=sitename%>-<%=siteurl%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="<%=sitedescription%>">
<meta name="keywords" content="<%=sitekeywords%>">
<link rel="stylesheet" href="buyok_shop.css" type="text/css">
</HEAD>
<CENTER>
<!--#include file="include/head.asp"-->
<table width=780 border=0 align=center cellpadding=0 cellspacing=0 bgcolor="#FFFFFF"  class="grayline" >
<tr> 
<td align="center" valign="top" width="185"> 

</td>
<td bgcolor="#CCCCCC" width="1"></td>
<td valign="top" bgcolor="#FfFfFf" width="570"> 
<TABLE border=0 width="100%" cellspacing="2" cellpadding="2" align=center>
<tr><td  height=30>目前位置：<a href=main.asp>首页</a> > <a href=user_center.asp>用户中心</a> > 站内消息</td></tr>
<tr><td width="570"  align="center" height="1" background="images/small/bgline.gif"></td></tr></table>
<br>
<%
sql="select * from sms where del='0' and (name='ALL' or name='"&request.cookies("cnhww")("username")&"') order by name asc, riqi desc"
set rs=Server.CreateObject("ADODB.recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then 
	response.write "<table width=90% border=0 align=center><tr><td><marquee width=100% height=10>暂时没有站内消息</marquee></td></tr></table>"
	else

do while not rs.eof and  pages>0
neirong=rs("neirong")
riqi=rs("riqi")
isnew=rs("isnew")
fname=rs("fname")
id=rs("id")
%>
<form action=my_msg.asp method=post>
<table width="90%" border="1" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" align=center bordercolordark="#FFFFFF" style="word-break:break-all">
<tr>
<td onMouseOver="bgColor='#e7e7e7'" onMouseOut="bgColor='#FFFFFF'">
<table border=0 width=100%>
<tr><td width=150 rowspan=2>
<a href='mymsg_hand.asp'><img alt="回复此消息" src=images/small/m_replyp.gif border=0></a> &nbsp;  
<span  style="cursor:hand" onclick="{if(confirm('提示：消息删除后不可恢复，\n\n您确实删除这条消息吗？ ')){location.href='my_msg.asp?delid=<%=id%>';}}"><img ALT="删除此消息" src=images/small/m_delete.gif border=0></span>&nbsp;&nbsp;
</td><td>

</td></tr>
<tr><td>发送人：<%=fname%>  &nbsp;  发送时间：<%=riqi%></td></tr>
<tr><td colspan=2><hr color="#FFFFFF" WIDTH=100%>
<%=replace(neirong,vbCRLF,"<BR>")%>
</td></tr>
</table>
</td></tr>
</table>
</form>
<%

rs.movenext
if rs.eof then exit do
loop
conn.execute "UPDATE sms SET isnew ='1' WHERE name ='"&request.cookies("cnhww")("username")&"'"
response.write "<table width=90% border=0 align=center><tr><td height=50 valign=top>"

response.write "</td></tr></table>"

end if
rs.close
set rs=nothing
%>
</td>
</TR>
</TABLE>
	<!--#include file="include/foot.asp"--></center>
</BODY>

<%
'分页
sub listPages() 
'if allpages <= 1 then exit sub 
if page = 1 then
response.write "<font color=darkgray>首页 前页</font>"
else
response.write "<a href="&request.ServerVariables("script_name")&"?page=1>首页</a> <a href="&request.ServerVariables("script_name")&"?page="&page-1&">前页</a>"
end if
if page = allpages then
response.write "<font color=darkgray> 下页 末页</font>"
else
response.write " <a href="&request.ServerVariables("script_name")&"?page="&page+1&">下页</a> <a href="&request.ServerVariables("script_name")&"?page="&allpages&">末页</a>"
end if
response.write " 第"&page&"页 共"&allpages&"页"
end sub

sub delmsg()
    Set rs = conn.Execute("select * from sms where id="&request("delid")) 
	if rs.eof and rs.bof then 
	response.write "<script language='javascript'>"
	response.write "alert('出错了，数据库操作失败！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif ucase(rs("name"))="ALL" then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您不能删除这条消息，可能这条消息属于公共消息！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    elseif Ucase(request.cookies("cnhww")("username"))<>Ucase(rs("name")) then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您不能删除这条消息，可能这条消息属于公共消息！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
    else
        conn.Execute ("update sms set del='1' where id="&request("delid")) 
	response.write "<script language='javascript'>"
	response.write "alert('操作成功，所选消息已被删除！');"
	response.write "location.href='user.asp?action=soumess';"
	response.write "</script>"
	response.end
    end if
end sub

conn.close
set conn=nothing
%>