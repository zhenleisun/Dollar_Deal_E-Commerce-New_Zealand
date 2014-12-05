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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<br><br><br><br><br><br><br><br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr> 
	<th background="../images/admin_bg_1.gif" class="tableHeaderText">商品类别转移</th>	
 </tr>
 <tr> 
	<td align="center" class="forumRowHighlight">转移小类的同时会转移小类下所有的商品；转移后需要修改小分类的排序。</td>	
 </tr>
  <form name="form1" method="post" action="savemoveclass.asp">
 <tr> 
      <td align="center" class="forumRowHighlight">选择要转移的小类：
<select name="nclassid" size="1" class="smallinput" >
	  <%set rs=server.CreateObject("adodb.recordset")
  		rs.Open "select nclassid,nclass from ssort	 order by nclassid",conn,1,1
		if rs.EOF and rs.BOF then
		response.Write "<option value=0>还没有分类</option>"
		else
		do while not rs.EOF
	%>
   <option value="<%=int(rs("nclassid"))%>"><%=trim(rs("nclass"))%></option>
   <%rs.MoveNext
   		loop
		rs.Close
		 set rs=nothing
		 end if%>
   </select>
    </td>
  </tr>
  <tr> 
      <td align="center" class="forumRowHighlight">请选择所属大类： 
        <select name="anclassid" size="1" class="smallinput" >
  <%set rs=server.CreateObject("adodb.recordset")
        rs.Open "select anclassid,anclass from bsort order by anclassidorder",conn,1,1
        if rs.eof and rs.bof then
        response.Write "<option value=0>还没有分类</option>"
        else
       do while not rs.eof
    %>
  <option value="<%=int(rs("anclassid"))%>"><%=trim(rs("anclass"))%></option>
  <%rs.movenext
       loop
       rs.close
       set rs=nothing
        end if%>
  </select>
  </td>
 </tr>
 <tr> 
   <td height="30" colspan="2" class="forumRowHighlight"><div align="center"><input type="submit" name="Submit" value="确定转移"></div></td>
  </tr>
 </form>
</table>
</body>
</html>
