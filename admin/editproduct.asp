<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
end if
dim bookid
bookid=request.QueryString("id")
if not isnumeric(bookid) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="edit.js" type="text/javascript"></script>
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>

<%


Dim htmlData

htmlData = Request.Form("bookcontent")

Function htmlspecialchars(str)
	str = Replace(str, "&", "&amp;")
	str = Replace(str, "<", "&lt;")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, """", "&quot;")
	htmlspecialchars = str
End Function
%>
<script type="text/javascript" charset="utf-8" src="../kindeditor.js"></script>
	<script type="text/javascript">
		KE.show({
			id : 'bookcontent',
			urlType : 'absolute',
			imageUploadJson : '../../asp/upload_json.asp',
			fileManagerJson : '../../asp/file_manager_json.asp',
			allowFileManager : true
			 
		});
	</script>

<body>

<%dim count
set rs=server.createobject("adodb.recordset")
rs.open "select * from ssort order by Nclassidorder ",conn,1,1%>
<script language = "JavaScript">
var onecount;
onecount=0;
subcat = new Array();
<%
   count = 0
   do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("Nclass"))%>","<%= rs("anclassid")%>","<%= rs("Nclassid")%>");
<%
        count = count + 1
        rs.movenext
        loop
        rs.close
%>
onecount=<%=count%>;
function changelocation(locationid)
    {
    document.myform.Nclassid.length = 0; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
             document.myform.Nclassid.options[document.myform.Nclassid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    
</script>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�޸���Ʒ</font></b></td>
</tr>
<tr> 
<form name="myform" method="post" action="saveaddproduct.asp?action=edit&id=<%=bookid%>">
<td> 
                                <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr > 
            <td width="30%" align="right" bgcolor="fbf4f4">ѡ����ࣺ</td>
            <td bgcolor="fbf4f4">���ࣺ 
              <%dim rs1
set rs1=server.CreateObject("adodb.recordset")
rs1.open "select * from products where bookid="&bookid,conn,1,1
    rs.open "select * from bsort order by anclassidorder",conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
%>
              <select name="anclassid" size="1" id="anclassid" onChange="changelocation(document.myform.anclassid.options[document.myform.anclassid.selectedIndex].value)">
                <% 
        do while not rs.eof
%>
                <option value="<%=rs("anclassid")%>" <%if rs1("anclassid")=rs("anclassid") then%>selected<%end if%>><%=trim(rs("anclass"))%></option>
                <%
        rs.movenext
        loop
		end if
        rs.close
%>
              </select>
              С�ࣺ 
              <select name="Nclassid">
                <%rs.open "select * from ssort where anclassid="&rs1("anclassid") ,conn,1,1
do while not rs.eof%>
                <option value="<%=rs("NclassID")%>" <%if rs1("nclassid")=rs("nclassid") then%>selected<%end if%>><%=rs("Nclass")%></option>
                <% rs.movenext
loop
        rs.close
        set rs = nothing
%>
              </select> </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��Ʒ���� </td>
            <td bgcolor="fbf4f4"> <input name="bookname" type="text" id="bookname" size="30" value="<%=trim(rs1("bookname"))%>"> 
            </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��Ʒ���</td>
            <td bgcolor="fbf4f4"><input name="grade" type="text" id="grade"  value="<%=trim(rs1("grade"))%>"></td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��ƷƷ��</td>
            <td bgcolor="fbf4f4"><input name="pingpai" type="text" id="pingpai" value="<%=trim(rs1("pingpai"))%>" size="20"> 
              <%
   set rs_s=server.createobject("adodb.recordset")
   rs_s.open "select * from brand order by pingpaiorder ",conn,1,1
   %>
              <select name="select" onChange="(document.myform.pingpai.value=this.options[this.selectedIndex].value)">option selected>��ѡ��Ʒ�� 
                <%
   while not rs_s.eof
   %>
                <option value="<%=rs_s("pingpainame")%>"><%=rs_s("pingpainame")%></option>
                <%
   rs_s.movenext
   wend
   rs_s.close
   set rs_s=nothing
   %>
              </select></td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4"> ��Ʒ��� </td>
            <td bgcolor="fbf4f4"> <input name="isbn" type="text" id="isbn" value="<%=trim(rs1("isbn"))%>"></td>
          </tr>
          <tr bgcolor="#fbf4f4" > 
            <td align="right">��Ʒ����</td>
            <td><input name="chima" type="text" id="chima" value="<%=trim(rs1("chima"))%>" size="30">
            </td>
          </tr>
          <tr bgcolor="#fbf4f4" > 
            <td align="right">��Ʒ��ɫ</td>
            <td><input name="yanse" type="text" id="yanse" value="<%=trim(rs1("yanse"))%>" size="30">
            </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4"> ��Ʒ��λ</td>
            <td bgcolor="fbf4f4"> <input name="bookchuban" type="text" id="bookchuban" size="8" value="<%=trim(rs1("bookchuban"))%>"> 
              <%
set rs_s=server.createobject("adodb.recordset")
rs_s.open "select * from unit order by danweiorder ",conn,1,1
%>
              <select name="select2" onChange="(document.myform.bookchuban.value=this.options[this.selectedIndex].value)">
                <option selected>��ѡ��λ</option>
                <%
while not rs_s.eof
%>
                <option value="<%=rs_s("danweiname")%>"><%=rs_s("danweiname")%></option>
                <%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
              </select> </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��Ʒ�۸� </td>
            <td bgcolor="fbf4f4"> �г��ۣ� 
              <input name="shichangjia" type="text" id="shichangjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("shichangjia")%>">
              ��Ա�ۣ� 
              <input name="huiyuanjia" type="text" id="huiyuanjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("huiyuanjia")%>">
              VIP�ۣ� 
              <input name="vipjia" type="text" id="vipjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))" value="<%=rs1("vipjia")%>"></td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��Ʒ��� </td>
            <td bgcolor="fbf4f4"> �� �棺 
              <input name="kucun" type="text" id="kucun" value="<%=rs1("kucun")%>" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))" >
              �����ۣ� 
              <input name="chengjiaocount" type="text" id="chengjiaocount" value="<%=rs1("chengjiaocount")%>" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))" readonly>
              �� �֣� 
              <input name="yeshu" type="text" id="yeshu" value="<%=rs1("yeshu")%>" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))"></td>
          </tr>
          <tr > 
            <td align="right" bgcolor="#fbf4f4">��ƷͼƬ </td>
            <td bgcolor="fbf4f4"> <input name="bookpic" type="text" id="bookpic" size="30" value="<%=trim(rs1("bookpic"))%>"> 
              &nbsp; <input type="button" name="Submit2" value="�ϴ�СͼƬ" onClick="window.open('../upload.asp?formname=myform&editname=bookpic&uppath=upfile/proimage&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"> 
            </td>
          </tr>
          <tr > 
            <td align="right" bgcolor="fbf4f4"></td>
            <td bgcolor="fbf4f4"><input name="zhuang" type="text" id="zhuang" size="30" value="<%=trim(rs1("zhuang"))%>"> 
              &nbsp; <input type="button" name="Submit2" value="�ϴ���ͼƬ" onClick="window.open('../upload.asp?formname=myform&editname=zhuang&uppath=upfile/proimage&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"> 
            </td>
          </tr>
          <tr > 
            <td rowspan="2" align="right" valign="top" bgcolor="fbf4f4">��ϸ˵��</td>
            <td bgcolor="fbf4f4"> <textarea id="bookcontent" name="bookcontent" cols="100" rows="8" style="width:600px;height:400px;visibility:hidden;"><%=htmlspecialchars(rs1("bookcontent"))%></textarea></td>
          </tr>
          <tr > 
            <td bgcolor="fbf4f4"><font color="ff0000">�ϴ�ͼƬ�� <strong>��� </strong>��Ȳ��ܳ��� 
              550 ���أ��ڲ�Ӱ�����ٵ�����½����ϴ�ͼƬ��С�� 100KB ���ڡ� </font></td>
          </tr>
          <tr > 
            <td align="right" bgcolor="fbf4f4"></td>
            <td height="30" bgcolor="fbf4f4"> <input name="newsbook" type="checkbox" id="newsbook" value="1" <%if rs1("newsbook")=1 then%>checked<%end if%>>
              ��Ʒ�� 
              <input name="bestbook" type="checkbox" id="bestbook" value="1" <%if rs1("bestbook")=1 then%>checked<%end if%>>
              �Ƽ��� 
              <input name="tejiabook" type="checkbox" id="tejiabook" value="1" <%if rs1("tejiabook")=1 then%>checked<%end if%>>
              �ؼ�<br>
              �� 
              <input type="submit" name="Submit" value="�޸ı���" OnClick=""></td>
          </tr>
        </table>
      </td>
    </form>
</tr>
</table>
<%rs1.close
set rs1=nothing
conn.close
set conn=nothing%>
<!--#include file="foot.asp"-->
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
function check()
{
   if(checkspace(document.myform.bookname.value)) {
	document.myform.bookname.focus();
    alert("��������Ʒ���ƣ�");
	return false;
  }
     if(checkspace(document.myform.shichangjia.value)) {
	document.myform.shichangjia.focus();
    alert("�������г��۸�");
	return false;
  }
     if(checkspace(document.myform.huiyuanjia.value)) {
	document.myform.huiyuanjia.focus();
    alert("�������Ա�۸�");
	return false;
  }
     if(checkspace(document.myform.vipjia.value)) {
	document.myform.vipjia.focus();
    alert("������VIP�۸�");
	return false;
  }

//-->
</script>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>
<%
function HTMLEncode(fString)
	fString = Replace(fString, "</P><P>", CHR(10) & CHR(10))
	fString = Replace(fString, "<BR>", CHR(10))
	HTMLEncode = fString
end function
%>
