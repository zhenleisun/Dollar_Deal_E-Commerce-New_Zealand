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
<title>品牌管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<meta http-equiv="目录类型" content="文本/html; 字符集=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<center>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
    <form method="post" action="SaveBrand.asp?action=update">
      <td class="forumRowHighlight" colspan="5" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">品牌设置与修改</font></b></td>	
      </tr>
      <tr align=center height=25> 
        <td class="forumRowHighlight">序号</td>
        <td class="forumRowHighlight">品牌名称</td>
        <td class="forumRowHighlight">是否推荐</td>
        <td class="forumRowHighlight">排序(必填项，数字)</td>
        <td class="forumRowHighlight">删除</td>
      </tr>
<%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open "SELECT  * from brand  order by pingpaiorder",conn,1,1
if rs.recordcount=0 then 
%>
      <tr> 
        <td colspan="8" height=25 align="center" width="100%" class="forumRowHighlight"> 还没有添加品牌 </td>
      </tr>
<%
else
    do while not rs.eof
%>
      <tr height=25> 
        <td align=center class="forumRowHighlight"><%=rs("ID")%></td>
        <td align=center class="forumRowHighlight"> 
          <input type="hidden" name="pingpaiid" value="<%=rs("id")%>">
          <input class=text type="text" name="pingpainame" size="10" value="<%=trim(rs("pingpainame"))%>">
        </td>
        <td align=center class="forumRowHighlight"> 
          <select class=text name="tuijian" size="1">
            <option value="1" <%if rs("tuijian")=1 then %>selected<%end if%>>推荐</option>
            <option value="0" <%if rs("tuijian")=0 then %>selected<%end if%>>不推荐</option>
          </select>
        </td>
        <td align=center class="forumRowHighlight"> 
          <input class=text type="text" name="pingpaiorder" size="10" value="<%=rs("pingpaiorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" >
        </td>
        <td align=center class="forumRowHighlight"> 
          <a href="KillBrand.asp?pingpaiid=<%=rs("id")%>">删除</a> 
        </td>
      </tr>
      <%
    rs.MoveNext
    Loop
%>
      <tr> 
        <td colspan="8" height=25 align="center" width="100%" class="forumRowHighlight"> 
          <input  type="submit" name="Submit2" value="保存修改" style="font-family: 宋体; font-size: 9pt" >
          &nbsp; 
          <input type="reset" value="重新设定" style="font-family: 宋体; font-size: 9pt" name="重置" >
        </td>
      </tr>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
    </form>
    <form method="post" action="SaveBrand.asp?action=add">
      <tr height=25> 
        <td align=center class="forumRowHighlight">添加品牌 
        </td>
        <td align=center class="forumRowHighlight"> 
          <input class=text type="text" name="pingpainame" size="10">
        </td>
        <td align=center class="forumRowHighlight"> 
          <select class=text name="tuijian" size="1">
            <option value="1" >推荐</option>
            <option value="0" >不推荐</option>
          </select>
        </td>
        <td align=center class="forumRowHighlight"> 
          <input class=text type="text" name="pingpaiorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();">
        </td>
        <td align=center colspan="2" class="forumRowHighlight"> 
          <input class="button" type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt">
        </td>
      </tr>
    </form>
  </table>
</center>
</body>
</html>