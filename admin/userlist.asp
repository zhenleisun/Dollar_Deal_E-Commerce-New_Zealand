<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<script language=JavaScript>
<%dim sql,i,j
	set rs_s=server.createobject("adodb.recordset")
	sql="select * from province order by shengorder"
	rs_s.open sql,conn,1,1
%>
	var selects=[];
	selects['xxx']=new Array(new Option('��ѡ����С���','xxx'));
<%
	for i=1 to rs_s.recordcount
%>
	selects['<%=rs_s("ShengNo")%>']=new Array(
<%
	set rs_s1=server.createobject("adodb.recordset")
	sql="select * from city where shengid="&rs_s("id")&" order by shiorder"
	rs_s1.open sql,conn,1,1
	if rs_s1.recordcount>0 then 
		for j=1 to rs_s1.recordcount
		if j=rs_s1.recordcount then 
%>
		new Option('<%=trim(rs_s1("shiname"))%>','<%=trim(rs_s1("shiNo"))%>'));
<%		else
%>
		new Option('<%=trim(rs_s1("shiname"))%>','<%=trim(rs_s1("shiNo"))%>'),
<%
		end if
		rs_s1.movenext
		next
	else 
%>
		new Option('','0'));
<%
	end if
	rs_s1.close
	set rs_s1=nothing
	rs_s.movenext
	next
rs_s.close
set rs_s=nothing
%>
	function chsel(){
		with (document.form1){
			if(szSheng.value) {
				szShi.options.length=0;
				for(var i=0;i<selects[szSheng.value].length;i++){
					szShi.add(selects[szSheng.value][i]);
				}
			}
		}
	}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<body>

<%dim userid
		userid=request.querystring("id")
		if userid<>"" then
if not isnumeric(userid) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from [user] where userid="&userid ,conn,1,1%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
  <form name="form1" method="post" action="saveuser.asp?action=save&id=<%=userid%>">
    <tr> 
      <td colspan="5" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�û���ϸ����</font></b></td>
    </tr>
    <tr> 
      <td width="40%"  style="PADDING-LEFT: 8px"><div align="right">�û����ƣ�</div></td>
      <td width="11%"  style="PADDING-LEFT: 8px"> <input name="username" type="text" id="username" size="12" value="<%=trim(rs("username"))%>" readonly></td>
      <td class="forumRowHighlight" width="65%">�û����� 
        <select name="reglx">
          <option value="1" <%if rs("reglx")=1 then%>selected<%end if%>>��ͨ��Ա</option>
          <option value="2" <%if rs("reglx")=2 then%>selected<%end if%>>VIP ��Ա</option>
        </select>
        VIP ���� 
        <input type="text" name="vipdate" size="10" value="<%=rs("vipdate")%>">
        2003-09-06</td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">��½���룺</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="userpassword" type="text" id="userpassword" size="12"> 
        <font color=#FF0000>Ϊ��<font color=#FF0000>���벻��</font>!</font></td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�������ʣ�</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="quesion" type="text" id="quesion" value="<%=trim(rs("quesion"))%>"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">����𰸣�</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="answer" type="text" id="answer"> 
        <font color=#FF0000> Ϊ��<font color=#FF0000>����𰸲���</font>!</font></td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">��ʵ������</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="userzhenshiname" type="text" id="userzhenshiname" size="12" value="<%=trim(rs("userzhenshiname"))%>"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�����ʼ���</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="useremail" type="text" id="useremail" value="<%=trim(rs("useremail"))%>"> 
        <%if rs("ifgongkai")=1 then %>
        (���������ַ) 
        <%else%>
        (�����������ַ) 
        <%end if%>
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">���֤���룺</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name=sfz type=text value="<%=trim(rs("sfz"))%>" maxlength="18" onKeyPress="event.returnValue=IsDigit();"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�� ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input type="radio" name="shousex" value="1" <%if rs("sex")=1 then%>checked<%end if%>>
        �С� 
        <input type="radio" name="shousex" value="2" <%if rs("sex")=2 then%>checked<%end if%>>
        Ů�� 
        <input type="radio" name="shousex" value="0" <%if rs("sex")=0 then%>checked<%end if%>>
        ����</td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�� �䣺</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name=nianling type=text value="<%=trim(rs("nianling"))%>" size="4" maxlength="2" onKeyPress="event.returnValue=IsDigit();"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">����ʡ/�У�</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <select size="1" name="szSheng" onChange=chsel()>
          <option value="xxx" selected>��ѡ��ʡ�ݡ���</option>
          <%dim tmpShengid
tmpShengid=0
set rs_s=server.createobject("adodb.recordset")
sql="select * from province  order by shengorder"
rs_s.open sql,conn,1,1
while not rs_s.eof
if rs("szSheng")=rs_s("ShengNo") then
tmpShengid=rs_s("id")
%>
          <option value="<%=rs_s("ShengNo")%>" selected ><%=trim(rs_s("ShengName"))%></option>
          <%
     else
%>
          <option value="<%=rs_s("ShengNo")%>" ><%=trim(rs_s("ShengName"))%></option>
          <%
end if
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
        </select> <select size="1" name="szShi">
          <%
set rs_s=server.createobject("adodb.recordset")
sql="select * from city where shengid="&tmpShengid&" order by shiorder"
rs_s.open sql,conn,1,1
while not rs_s.eof
%>
          <option value="<%=rs_s("ShiNo")%>" <%if rs("szShi")=rs_s("ShiNo") then%>selected<%end if%>><%=trim(rs_s("ShiName"))%></option>
          <%
rs_s.movenext
wend
rs_s.close
set rs_s=nothing
%>
        </select> </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">��ϸ��ַ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="shouhuodizhi" type="text" id="shouhuodizhi" size="40" value="<%=trim(rs("shouhuodizhi"))%>"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�� �ࣺ</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="youbian" type="text" id="youbian" size="12" value="<%=rs("youbian")%>" maxlength=6 onKeyPress="event.returnValue=IsDigit();"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">��ϵ�绰��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name="usertel" type="text" id="usertel" size="12" value="<%=trim(rs("usertel"))%>"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">Q Q��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name=QQ type=text value="<%=trim(rs("oicq"))%>" size="12" maxlength="12"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">������ҳ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name=homepage type=text value="<%=trim(rs("homepage"))%>" size="40"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�� �飺</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <textarea name="content" rows="6" cols="40"><%=trim(rs("content"))%></textarea> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">Ԥ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> <input name=yucun type=text value="<%=trim(rs("yucun"))%>" size="10" maxlength="8" onKeyPress="event.returnValue=IsDigit();"> 
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">�ͻ���ʽ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> 
        <%
	set rs_s=Server.createobject("adodb.recordset")
	rs_s.open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
	if rs_s.recordcount=0 then
	response.write "��δѡ��"
	else
	response.write rs_s("subject")
	end if
	rs_s.close
	set rs_s=nothing
	%>
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">֧����ʽ��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> 
        <%
	set rs_s=Server.createobject("adodb.recordset")
	rs_s.open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
	if rs_s.recordcount=0 then
	response.write "��δѡ��"
	else
	response.write rs_s("subject")
	end if
	rs_s.close
	set rs_s=nothing
	%>
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">ע��ʱ�䣺</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"><%=rs("adddate")%></td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">����½��</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"><%=rs("lastlogin")%></td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">��½������</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"><%=rs("logins")%> ��</td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px" width="40%"><div align="right">���¶�������</div></td>
      <td  style="PADDING-LEFT: 8px" colspan="3"> 
        <%dim rs2
			set rs2=server.CreateObject("adodb.recordset")
			rs2.open "select distinct(dingdan) from orders where username='"&trim(rs("username"))&"' and zhuangtai<=5",conn,1,1
			response.write rs2.recordcount&"�ʶ���"
			rs2.close
			set rs2=nothing
			%>
      </td>
    </tr>
    <tr> 
      <td  style="PADDING-LEFT: 8px"><div align="right">���Ҵ��û������ж�����</div></td>
      <td height="30" colspan="3"  style="PADDING-LEFT: 8px"> <select name="zhuangtai" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" ><base target=Right> 
          <option value="" selected>--ѡ���Ѷ״̬--</option>
          <option value="editorder.asp?zhuangtai=0&namekey=<%=trim(rs("username"))%>" >ȫ������״̬</option>
          <option value="editorder.asp?zhuangtai=1&namekey=<%=trim(rs("username"))%>" >δ���κδ���</option>
          <option value="editorder.asp?zhuangtai=2&namekey=<%=trim(rs("username"))%>" >�û��Ѿ�������</option>
          <option value="editorder.asp?zhuangtai=3&namekey=<%=trim(rs("username"))%>" >�������Ѿ��յ���</option>
          <option value="editorder.asp?zhuangtai=4&namekey=<%=trim(rs("username"))%>" >�������Ѿ�����</option>
          <option value="editorder.asp?zhuangtai=5&namekey=<%=trim(rs("username"))%>" >�û��Ѿ��յ���</option>
        </select> </td>
    </tr>
    <%rs.close
			set rs=nothing
			conn.close
			set conn=nothing%>
    <tr> 
      <td  style="PADDING-LEFT: 8px"></td>
      <td height="28" colspan="3" > <input type="submit" name="Submit" value="ȷ���ύ"> 
        &nbsp; <input type="button" name="Submit2" value="������һҳ" onClick='javascript:history.go(-1)'></td>
    </tr>
  </form>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
