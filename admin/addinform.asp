
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
function changechr(str) 
    changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
    changechr=replace(changechr,"'","&quot;")
    changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
end function
Content=Request.Form("Content")
 aa="http://"&Request.ServerVariables("server_name")&Request.ServerVariables("path_info")
 aa= mid(aa,1,len(aa)-21)		
content=replace(content,aa,"")
session("content")=content
if Content=""  then
%>
<script language=javascript>
history.back()
alert("请输入文章内容！")
</script>
<%
Response.End
end if
title=request.form("title")
if title="" then
%>
<script language=javascript>
history.back()
alert("请填写文章标题！")
</script>
<%
Response.End
end if
if len(title)>100 then
%>
<script language=javascript>
history.back()
alert("文章标题过长！")
</script>
<%
Response.End
end if
title=changechr(trim(request.form("title")))
if title_color="0" then
	title_color="#003399"
end if
title_size=Request.Form("title_size")
if title_size="0" then
	title_size="5"
end if
title_type=Request.Form("title_type")
if title_type="0" then
	title_type="b"
end if
title_face=Request.Form("title_face")
if title_face="0" then
	title_face="楷体_GB2312"
end if
set rs=server.createobject("adodb.recordset")
sql="select * FROM inform" 
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("content")=content
rs("UpdateTime")=now()
rs("titletype")=title_type
rs("titlesize")=title_size
rs("titlecolor")=title_color
rs("titleface")=title_face
rs.update
rs.Close
set rs=nothing
session("content")=""
conn.close
set conn=nothing
%>
<script language="javascript">
window.alert("文章添加成功!");
window.location.href='addinforms.asp'
</script>
