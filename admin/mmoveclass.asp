<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
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
	<th background="../images/admin_bg_1.gif" class="tableHeaderText">��Ʒ���ת��</th>	
 </tr>
 <tr> 
	<td align="center" class="forumRowHighlight">ת��С���ͬʱ��ת��С�������е���Ʒ��ת�ƺ���Ҫ�޸�С���������</td>	
 </tr>
  <form name="form1" method="post" action="savemoveclass.asp">
 <tr> 
      <td align="center" class="forumRowHighlight">ѡ��Ҫת�Ƶ�С�ࣺ
<select name="nclassid" size="1" class="smallinput" >
	  <%set rs=server.CreateObject("adodb.recordset")
  		rs.Open "select nclassid,nclass from ssort	 order by nclassid",conn,1,1
		if rs.EOF and rs.BOF then
		response.Write "<option value=0>��û�з���</option>"
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
      <td align="center" class="forumRowHighlight">��ѡ���������ࣺ 
        <select name="anclassid" size="1" class="smallinput" >
  <%set rs=server.CreateObject("adodb.recordset")
        rs.Open "select anclassid,anclass from bsort order by anclassidorder",conn,1,1
        if rs.eof and rs.bof then
        response.Write "<option value=0>��û�з���</option>"
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
   <td height="30" colspan="2" class="forumRowHighlight"><div align="center"><input type="submit" name="Submit" value="ȷ��ת��"></div></td>
  </tr>
 </form>
</table>
</body>
</html>
