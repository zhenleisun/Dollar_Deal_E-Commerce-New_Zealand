<!-- ��ʹ�õ�����Ȥ���Ϲ���ϵͳ��Ѱ汾����Ѱ�������ܾ���ʹ�ã����߼������ǲ��߱��ģ������Ҫ���Ͽ��꣬�Ƽ�ʹ�ùٷ���ʽ�棬�ٷ���վ:http://www.cnhww.com �ͷ�QQ��81447932  -->

<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")=2 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
<script language="JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<%
dim selectm,selectkey,selectid
selectkey=trim(request.form("selectkey"))
selectm=trim(request.form("selectm"))
action=trim(request("action"))
if selectkey="" then
selectkey=FormatSQL(request.QueryString("selectkey"))
end if
if selectm="" then
selectm=FormatSQL(request.QueryString("selectm"))
end if
selectid=request.form("selectid")

select case action
case "update"
j = request.form("updateid").count
    for i=1 to request.form("updateid").count
	updateid=replace(request.form("updateid")(i),"'","")
	strname=replace(request.form("strname")(i),"'","")
	price1=replace(request.form("price1")(i),"'","")
	price2=replace(request.form("price2")(i),"'","")
	price3=replace(request.form("price3")(i),"'","")
	price4=replace(request.form("price4")(i),"'","")
	price5=replace(request.form("price5")(i),"'","")
	
	stock=replace(request.form("stock")(i),"'","")
	stock2=replace(request.form("stock2")(i),"'","")
'	orderby=replace(request.form("orderby")(i),"'","")
	score=replace(request.form("score")(i),"'","")
'	viewnum=replace(request.form("viewnum")(i),"'","")
'	solded=replace(request.form("solded")(i),"'","")

	conn.execute("update  products set bookname='"&strname&"',grade='"&price1&"',isbn='"&price2&"',shichangjia="&price3&",huiyuanjia="&price4&",vipjia="&price5&",kucun="&stock&",bookchuban='"&stock2&"',yeshu="&score&" where bookid="&updateid)
    next
conn.close
set conn=nothing
response.Redirect "chkpro.asp"
response.End

end select


 

%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6699cc">
  <tr>  
    <td align="center" background="../images/admin_bg_1.gif" height="25"><b><font color="#ffffff">��Ʒ�����޸Ĺ���</font></b> 
    </td>
  </tr>
  <tr> 
    <form name="chkform" method="post" action="">
      <td bgcolor="#FFFFFF"><br>
        <%
				Const MaxPerPage=20
   				dim totalPut   
   				dim CurrentPage
   				dim TotalPages
   				dim j
   				dim sql
    				if Not isempty(SafeRequest("page",1)) then
      				currentPage=Cint(SafeRequest("page",1))
   				else
      				currentPage=1
   				end if 
			set rs=server.CreateObject("adodb.recordset")
			select case selectm
			case ""
            rs.open "select * from  [products]  order by bookid desc",conn,1,1

		  end select
				
  				if rs.eof And rs.bof then
       				Response.Write "<p align='center' class='contents'> ���ݿ�����ʱ�����ݣ�</p>"
   
		else
	  		rs.PageSize =20 'ÿҳ��¼����
			iCount=rs.RecordCount '��¼����
			iPageSize=rs.PageSize
    		maxpage=rs.PageCount 
    		page=request("page")
    if Not IsNumeric(page) or page="" then
        page=1
    else
        page=cint(page)
    end if
    if page<1 then
        page=1
    elseif  page>maxpage then
        page=maxpage
    end if
    rs.AbsolutePage=Page
	if page=maxpage then
		x=iCount-(maxpage-1)*iPageSize
	else
		x=iPageSize
	end if
		%><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
       <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#A4B6D7">
          <tr bgcolor="#A4B6D7" height="25"> 
            <td width="8%" align="center">ID</td>
            <td width="17%" align="center">��Ʒ����</td>
            <td width="7%" align="center">��Ʒ���</td>
            <td width="10%" align="center">��Ʒ���</td>
            <td width="8%" align="center">�г���</td>
            <td width="8%" align="center">��Ա��</td>
            <td width="8%" align="center">VIP��</td>
            <td width="8%" align="center">���</td>
            <td width="8%" align="center">��Ʒ��λ</td>
            <td width="8%" align="center">����</td>
            <td width="10%" align="center">�޸�</td>

          </tr>
          <%
		  do while not rs.eof 
		   iii = iii +1
		  if iii mod  2 = 0 then  %>
          <tr bgcolor="#FFFFFF" align="center" height="28"> 
            <% else %>
          <tr bgcolor="#EFEFEF" height="28"> 
            <% end if %>
            <td align="center"><a href=editproduct.asp?id=<%=rs("bookid")%>><%=rs("bookid")%></a></td>
            <td align="left"><input type="text" name="strname" size="20" value="<% = rs("bookname") %>"></td>
            <td align="center">
<input type="text" name="price1" size="6" value="<%=rs("grade") %>"></td>
            <td align="center">
<input type="text" name="price2" size="6" value="<%=rs("isbn") %>"></td>
            <td align="center">
<input type="text" name="price3" size="6" value="<%=FormatNum(rs("shichangjia"),2)%>"></td>
            <td align="center">
<input type="text" name="price4" size="6" value="<%=FormatNum(rs("huiyuanjia"),2)%>"></td>
            <td align="center">
<input type="text" name="price5" size="6" value="<%=FormatNum(rs("vipjia"),2)%>"></td>
            <td align="center">
<input type="text" name="stock" size="3" value="<%= rs("kucun") %>" ONKEYPRESS="event.returnValue=IsDigit();"></td>
            <td align="center">
<input type="text" name="stock2" size="3" value="<%= rs("bookchuban") %>" onKeyPress="event.returnValue=IsDigit();"></td>
            <td align="center">
<input type="text" name="score" size="3" value="<%= rs("yeshu") %>" ONKEYPRESS="event.returnValue=IsDigit();"></td>
            <td align="center"><a href=editproduct.asp?id=<%=rs("bookid")%>>�޸���Ʒ</a> 
            </td>
            <input type="hidden" name="updateid" value="<%=rs("bookid")%>">
          </tr>
          <%i=i+1
			if i>=MaxPerPage then Exit Do
			rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
          <tr align="center" bgcolor="#FFFFFF"> 
            <td height="30" colspan="14"> 
              <INPUT type="hidden" name="action" value=""> 
              <INPUT type="button" value="������Ʒ��" onclick="document.chkform.action.value='update'; document.chkform.submit();">
              �������Ʒ����������޸���Ʒ����! &nbsp;&nbsp; </td>
          </tr>
        </table>
 <%
		call PageControl(iCount,maxpage,page,"border=0 align=center","<p align=center>")
		end if
		rs.close
		set rs=nothing
Sub PageControl(iCount,pagecount,page,table_style,font_style)
'������һҳ��һҳ����
    Dim query, a, x, temp
    action = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("SCRIPT_NAME")
    query = Split(Request.ServerVariables("QUERY_STRING"), "&")
    For Each x In query
        a = Split(x, "=")
        If StrComp(a(0), "page", vbTextCompare) <> 0 Then
            temp = temp & a(0) & "=" & a(1) & "&"
        End If
    Next
    Response.Write("<table width=100% border=0 cellpadding=0 cellspacing=0 >" & vbCrLf )        
   Response.Write("<TD align=center height=40>" & vbCrLf )
    Response.Write(font_style & vbCrLf )    
    if page<=1 then
        Response.Write ("�� ҳ " & vbCrLf)        
        Response.Write ("��һҳ " & vbCrLf)
    else        
        Response.Write("<A HREF=" & action & "?" & temp & "Page=1>�� ҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page-1) & ">��һҳ</A> " & vbCrLf)
    end if
    if page>=pagecount then
        Response.Write ("��һҳ " & vbCrLf)
        Response.Write ("β ҳ " & vbCrLf)            
    else
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & (Page+1) & ">��һҳ</A> " & vbCrLf)
        Response.Write("<A HREF=" & action & "?" & temp & "Page=" & pagecount & ">β ҳ</A> " & vbCrLf)            
    end if
    Response.Write(" ҳ�Σ�" & page & "/" & pageCount & "ҳ" &  vbCrLf)
    Response.Write(" ����" & iCount & "����Ʒ" &  vbCrLf)
    
    Response.Write("</TD>" & vbCrLf )                
    
    Response.Write("</table>" & vbCrLf )        
End Sub
%></td>
 
  
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#6699cc">
  <tr> 
    
  </tr>
 
</table>

<!--#include file="foot.asp"-->
<script language=javascript>

function mm()
{
   var a = document.getElementsByTagName("input");
   if(a[0].checked==true){
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = false;
   }
   else
   {
   for (var i=0; i<a.length; i++)
      if (a[i].type == "checkbox") a[i].checked = true;
   }
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<%
function FormatNum(num,n)
if num<1 then
num="0"&cstr(FormatNumber(num,n))
else
num=cstr(FormatNumber(num,n))
end if
FormatNum=replace(num,",","")
end function
%>
