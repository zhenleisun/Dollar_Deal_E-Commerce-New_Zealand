<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
set rs=server.createobject("adodb.recordset")
sql="select bookid,bookname,pingpai,bookchuban,yeshu,isbn,bookpic,shichangjia,huiyuanjia,vipjia,bestbook,tejiabook,newsbook,kucun,adddate,grade from  products order by adddate desc "
rs.open sql,conn,1,1
%>
<br>
<td height="107" bgcolor="#DEE7FF"> 
<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#6396D6">
    <tr align="center" bgcolor="#C7D3E6"> 
      <td width="11%" bgcolor="#7B9AE7" height="25"><font color="#FFFFFF">商品ID</font></td>
      <td width="20%" bgcolor="#7B9AE7"><font color="#FFFFFF">商品名称</font></td>
      <td width="21%" bgcolor="#7B9AE7"><font color="#FFFFFF">所属品牌</font></td>
      <%
if not rs.eof then
%>
      <td width="17%" bgcolor="#7B9AE7"><font color="#FFFFFF">商品编号</font></td>
      <%
else
%>
      <td width="17%" bgcolor="#7B9AE7"><font color="#FFFFFF">商品单位</font></td>
      <%
end if
%>
      <td width="14%" bgcolor="#7B9AE7"><font color="#FFFFFF">添加日期</font></td>
    </tr>
    <%
while not rs.eof
%>
    <tr bgcolor="#FFFFFF"> 
      <td width="11%" align="center" bgcolor="#DEE7FF"><%=rs(0)%></td>
      <td width="20%" height="22" bgcolor="#DEE7FF"><a href="editproduct.asp?id=<%=rs(0)%>"><%=rs(1)%></a></td>
      <td width="21%" height="22" align="center" bgcolor="#DEE7FF"><%=rs(2)%></td>
      <td width="17%" align="center" bgcolor="#DEE7FF"><%=rs(15)%></td>
      <td width="14%" align="center" bgcolor="#DEE7FF"><%=rs(14)%></td>
    </tr>
    <%
rs.movenext
wend
set rs=nothing
%>
    <tr bgcolor="#FFFFFF"> 
      <td height="22" colspan="6" bgcolor="#DEE7FF"> <div align="center"> 
          <input type="button" name="Submit4" value="打 印" onClick="javascript:window.print()">
          <input type="button" name="Submit" value="导入EXCEL(点此后会一闪同时本按钮右侧会有一个链接)" onClick="javascript:openexcel();">
          <font color="red"><span id="excelfile" ></span></font></div></td>
    </tr>
  </table>
</td>
</tr>
</table>
<br>
</body>
</html>
<script>

function openexcel(){
    window.open("toexcel.asp?sql=<%=sql%>");
}
</script>