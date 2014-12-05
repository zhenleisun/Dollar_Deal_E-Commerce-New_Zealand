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
<%dim anclassid,anclass,paixu
anclass=request.QueryString("nclass")
anclassid=request.QueryString("id")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">商品小类管理</font></b></td>
</tr>
<tr> 
<td width="30%" valign="top" align="center" ><br>
<select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
<option >选择商品分类</option>
                                  <%set rs=server.createobject("adodb.recordset")
		rs.Open "select * from bsort order by anclassidorder",conn,1,1
		do while not rs.eof %>
                                  <option value="managessort.asp?id=<%=rs("anclassid")%>&anclass=<%=rs("anclass")%>" <%if rs("anclassid")=cint(request.QueryString("id")) then %>selected<%end if%>><%=trim(rs("anclass"))%></option>
                                  <%rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
                                </select><br>
                                <%if request.QueryString("id")<>"" then
        response.Write "当前查讯："&request.QueryString("anclass")
        end if%>
</td>
<td width="70%" > 
<table width="90%" align="center" border="0" cellpadding="1" cellspacing="2">
<tr> 
<td width="50%" align="center">分类名称</td>
<td width="20%" align="center">分类排序</td>
<td width="30%" align="center">确定操作</td>
</tr>
                                <%
        if anclassid="" then
        response.Write "<div align=center><font color=red>请选择左测的分类</font></div>"
        else
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * FROM ssort where anclassid="&anclassid&" order by nclassidorder",conn,1,1
         if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有分类</font></center>"
		  paixu=0
		  else
         do while not rs.EOF
         %>
<form name="form1" method="post" action="savessort.asp?action=edit&id=<%=rs("nclassid")%>&anclass=<%=request.QueryString("anclass")%>">
<tr> 
<td align="center">
<input name="nclass" type="text" id="nclass" size="16" value="<%=trim(rs("nclass"))%>">
<input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>" id="Hidden1">
</td>
<td align="center">
<input name="nclassidorder" type="text" id="nclassidorder" size="4" value="<%=int(rs("nclassidorder"))%>">
</td>
<td align="center">
<input type="submit" name="Submit" value="修 改">&nbsp;
<a href="savessort.asp?id=<%=int(rs("nclassid"))%>&action=del&anclassid=<%=request.QueryString("id")%>&anclass=<%=request.QueryString("anclass")%>" onClick="return confirm('您确定进行删除操作吗？')"><font color=red>删除</font></a> 
</td>
</tr>
</form>
		<%rs.movenext
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        end if
		%>
</table>
</td>
</tr>
</table><br>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加商品分类</font></b></td>
</tr>
<tr> 
<td width="30%" valign="top" align="center" ><br>
当前分类：<%=request.QueryString("anclass")%></div>
</td>
<td width="70%" > 
<table width="90%" align="center" border="0" cellpadding="1" cellspacing="2">
<tr> 
<td width="50%" align="center">分类名称</div></td>
<td width="20%" align="center">分类排序</div></td>
<td width="30%" align="center">确定操作</div></td>
</tr>
<form name="form2" method="post" action="savessort.asp?action=add&anclass=<%=request.QueryString("anclass")%>">
<tr> 
<td align="center"> 
<input name="nclass2" type="text" id="nclass22" size="16">
<input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>">
</td>
<td align="center"> 
<input name="nclassidorder2" type="text" id="nclassidorder22" size="4" value="<%=paixu+1%>">
</td>
<td align="center"> 
<input type="submit" name="Submit2" value="添加分类">
</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
