<%dim bookid,action
bookid=request.QueryString("id")
action=request.QueryString("action")
if action="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from review",conn,1,3
rs.addnew
rs("bookid")=request("bookid")
rs("pingji")=request("pingji")
rs("pinglunname")=HTMLEncode2(trim(request("pinglunname")))
rs("pingluntitle")=HTMLEncode2(trim(request("pingluntitle")))
rs("pingluncontent")=HTMLEncode2(trim(request("pingluncontent")))
rs("ip")=Request.servervariables("REMOTE_ADDR")
rs("pinglundate")=now()
rs("shenhe")=1
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
response.Write "<script language=javascript>alert('Your comment has been successfully submitted£¡');history.go(-1);</script>"
response.End
end if
%><table width="95%" align="center" border="0" cellpadding="2" cellspacing="1" bgcolor="#cccccc">
<tr bgcolor="#ffffff"> 
<form name="pinglunform" method="post" action="products.asp?action=save&id=<%=id%>">
<td> 
<table width="95%" align="center" border="0" cellpadding="2" cellspacing="1">
<tr> 
<td width="10%">Name£º</td>
<td width="90%">
<input class="wenbenkuang" name="pinglunname" type="text" id="pinglunname" maxlength="18">
<input type="radio" name="pingji" value="1" checked>¡î
<input type="radio" name="pingji" value="2">¡î¡î
<input type="radio" name="pingji" value="3">¡î¡î¡î
<input type="radio" name="pingji" value="4">¡î¡î¡î¡î
<input type="radio" name="pingji" value="5">¡î¡î¡î¡î¡î
</td>
</tr>
<tr> 
<td width="10%">Title£º</td>
<td width="90%">
<input class="wenbenkuang" name="pingluntitle" type="text" id="pingluntitle" size="70">
</td>
</tr>
<tr> 
<td valign="top">Content£º</td>
<td>
<textarea class="wenbenkuang" name="pingluncontent" cols="70" rows="3" id="pingluncontent" style="BORDER-RIGHT: #ffffff 1px groove; BORDER-TOP: BORDER-LEFT: COLOR: #333333; BORDER-BOTTOM: HEIGHT: 18px; BACKGROUND-COLOR:"; ";";";"></textarea>
</td>
</tr>
<tr>
<td></td>
<td>
<input class="go-wenbenkuang" name="submit" value="Submit" type="submit" onClick="return check();">
</td>
</tr>
</table>
</td>
</form>
</tr>
</table>
<%function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "<I></I>&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
	HTMLEncode2 = fString
end function%>
<script LANGUAGE="javascript">
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.pinglunform.pinglunname.value)) {
	document.pinglunform.pinglunname.focus();
    alert("Please fill in your name£¡");
	return false;
  }
  if(checkspace(document.pinglunform.pingluntitle.value)) {
	document.pinglunform.pingluntitle.focus();
    alert("Please fill in the title£¡");
	return false;
  }
  if(checkspace(document.pinglunform.pingluncontent.value)) {
	document.pinglunform.pingluncontent.focus();
    alert("Please fill in the comment text£¡");
	return false;
  }
	  }
//-->
</script>
