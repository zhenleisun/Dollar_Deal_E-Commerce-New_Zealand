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
<%dim count
set rs=server.createobject("adodb.recordset")
rs.open "select * from ssort order by nclassidorder ",conn,1,1

%>
<script language = "JavaScript">
var onecount;
var threecount;
onecount=0;
threecount=0
subcat3=new Array();
<%
count = 0

%>

threecount=<%=count%>;
subcat = new Array();
<%
count = 0
do while not rs.eof 
%>
subcat[<%=count%>] = new Array("<%= trim(rs("nclass"))%>","<%= rs("anclassid")%>","<%= rs("nclassid")%>");
<%
count = count + 1
rs.movenext
loop
rs.close
%>	
onecount=<%=count%>;
function changelocation(locationid)
    {
    document.myform.nclassid.length = 0; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
             document.myform.nclassid.options[document.myform.nclassid.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
     changelocation3(document.myform.nclassid.options[document.myform.nclassid.selectedIndex].value)
    }
function changelocation3(locationid)
    {
    document.myform.xclassid.length = 0; 
    var locationid=locationid;
    var i;
    for (i=0;i < threecount; i++)
        {
            if (subcat3[i][1] == locationid)
            { 
             document.myform.xclassid.options[document.myform.xclassid.length] = new Option(subcat3[i][0], subcat3[i][2]);
            }
        }
    } 
	function addpl(spnum)
{

var spnum1=""

if(spnum!=""){

	pltjsp.style.display=""

	for (i=2;i<=spnum;i++)
	{
	spnum1=spnum1+"&nbsp;批量商品名称"+i+"：<input name=\"bookname1\" type=\"text\"  size=\"30\">"+"<br>"
	}
	pltjsp1.innerHTML=spnum1
	}

}
//-->
</script>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
  <tr > 
    <td height="12" colspan="2" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">添加新商品</font></b> 
    </td>
  </tr>
  <tr> 
    <td align="left"><div align="center">批量添加商品 
        <input type="text" id="plnum" size="5" maxlength="4">
        条 
        <input type="button" value="增加" onClick="addpl(plnum.value)">
        请输入批量添加商品的个数。</div></td>
    <td align="left">&nbsp;</td>
  </tr>
  <tr > 
    <td height="12" colspan="2" align="center">&nbsp;</td>
  </tr>
  <tr > 
    <form name="myform" method="post" action="saveaddproduct.asp?action=Batch" OnSubmit="return checkkk()" >
      <td colspan="2"> <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#fbf4f4">
          <tr > 
            <td width="30%" align="right">选择商品的分类：</td>
            <td width="70%"> 大类： 
              <%
rs.open "select * from bsort order by anclassidorder",conn,1,1
if rs.eof and rs.bof then
response.write "请先添加分类"
response.end
else
%>
              <select name="anclassid" size="1" id="select2" onChange="changelocation(document.myform.anclassid.options[document.myform.anclassid.selectedIndex].value)">
                <%
dim selclass
selclass=rs("anclassid")
do while not rs.eof
response.write  "<option value="&rs("anclassid")
if anclassid=rs("anclassid") then response.write " selected "
response.write ">"&trim(rs("anclass"))&"</option>"    
rs.movenext
loop
end if
rs.close
%>
              </select>
              小类： 
              <select name="nclassid" size="1" id="nclassid" onChange="changelocation3(document.myform.nclassid.options[document.myform.nclassid.selectedIndex].value)">
                <%
rs.open "select * from ssort where anclassid="&selclass,conn,1,1
if not(rs.eof and rs.bof) then
%>
                <% 
dim selclass3
selclass3=rs("nclassid")
do while not rs.eof
response.write  "<option value="&rs("nclassid")
if nclassid=rs("nclassid") then response.write " selected "
response.write ">"&trim(rs("nclass"))&"</option>"
rs.movenext
loop
end if
rs.close
set rs = nothing
%>
              </select>
            </td>
          </tr>
          <tr > 
            <td align="right">商品名称 </td>
            <td><input name="bookname" type="text" id="bookname" size="30"> </td>
          </tr>
          <tr id="pltjsp" style="DISPLAY: none"> 
            <td align="center">批量商品名称</td>
            <td colspan="3"><font id="pltjsp1">&nbsp;</font></td>
  </tr>
          <tr > 
            <td align="right">商品编号</td>
            <td><input name="grade" type="text" id="grade" ></td>
          </tr>
          <tr > 
            <td align="right">所属品牌</td>
            <td><input name="pingpai" type="text" id="pingpai" size="20"> 
              <%
   set rs_s=server.createobject("adodb.recordset")
   rs_s.open "select * from brand order by pingpaiorder ",conn,1,1
   %>
              <select name="select2" onchange="(document.myform.pingpai.value=this.options[this.selectedIndex].value)">option selected>请选择品牌 
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
            <td align="right"> 商品规格 </td>
            <td><input name="isbn" type="text" id="isbn"></td>
          </tr>
          <tr > 
            <td align="right">商品尺码</td>
            <td><input name="chima" type="text" id="chima" size="30">
              (请用 / 作为不同颜色的分隔,可为空)</td>
          </tr>
          <tr > 
            <td align="right">商品颜色</td>
            <td><input name="yanse" type="text" id="yanse" size="30">
              (请用 / 作为不同颜色的分隔,可为空)</td>
          </tr>
          <tr > 
            <td align="right"> 商品单位</td>
            <td> <input name="bookchuban" type="text" id="bookchuban" size="7"> 
              <%
set rs_s=server.createobject("adodb.recordset")
rs_s.open "select * from unit order by danweiorder ",conn,1,1
%>
              <select name="select" onChange="(document.myform.bookchuban.value=this.options[this.selectedIndex].value)">
                <option selected>请选择单位</option>
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
            <td align="right">商品价格 </td>
            <td> 市场价： 
              <input name="shichangjia" type="text" id="shichangjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
              会员价： 
              <input name="huiyuanjia" type="text" id="huiyuanjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))">
              VIP价： 
              <input name="vipjia" type="text" id="vipjia" size="6" onKeyPress	= "return regInput(this,	/^\d*\.?\d{0,2}$/,		String.fromCharCode(event.keyCode))" 
		onpaste		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^\d*\.?\d{0,2}$/,		event.dataTransfer.getData('Text'))"></td>
          </tr>
          <tr > 
            <td align="right">商品库存 </td>
            <td> 库存数： 
              <input name="kucun" type="text" id="kucun" value="0" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))" >
              已销售： 
              <input name="chengjiaocount" type="text" id="chengjiaocount" value="0" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))" readonly>
              积 分： 
              <input name="yeshu" type="text" id="yeshu" value="0" size="6" onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))" >	
            </td>
          </tr>
          <tr > 
            <td align="right">商品图片 </td>
            <td> <input name="bookpic" type="text" id="bookpic" size="30"> &nbsp; 
              <input type="button" name="Submit2" value="上传小图片" onClick="window.open('../upload.asp?formname=myform&editname=bookpic&uppath=upfile/proimage&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td></td>
            <td> <input name="zhuang" type="text" id="zhuang" size="30"> &nbsp; 
              <input type="button" name="Submit22" value="上传大图片" onClick="window.open('../upload.asp?formname=myform&editname=zhuang&uppath=upfile/proimage&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="right">META_DESCRIPTION：</td>
            <td><textarea name="metad" cols="66" rows="2" id="metad"></textarea> 
              <br>
              可留空，系统自动采用网站设置中的META_DESCRIPTION信息 </td>
          </tr>
          <tr > 
            <td align="right">META_KEYWORDS：</td>
            <td><textarea name="metak" cols="66" rows="2" id="metak"></textarea> 
              <br>
              可留空，系统自动采用网站设置中的META_KEYWORDS信息 </td>
          </tr>
          <tr > 
            <td rowspan="2" align="right" valign="top"> 详细说明</td>
            <td><input name="bookcontent" type="hidden" id="bookcontent" value=""> 
              <IFRAME ID="bookcontent" SRC="editubb/ewebeditor.asp?id=bookcontent&style=standard" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME></td>
          </tr>
          <tr > 
            <td><font color="ff0000">上传图片的 <strong>最大 </strong>宽度不能超过 550 像素，在不影响网速的情况下建议上传图片大小在 
              100KB 以内。 </font></td>
          </tr>
          <tr > 
            <td></td>
            <td height="30"> <input name="newsbook" type="checkbox" id="newsbook" value="1">
              新 品　 
              <input name="bestbook" type="checkbox" id="bestbook" value="1">
              热 销　 
              <input name="tejiabook" type="checkbox" id="tejiabook" value="1">
              特 价</td>
          </tr>
          <tr > 
            <td></td>
            <td height="30"> <input type="submit" name="Submit" value="提交保存" onClick="document.myform.bookcontent.value = frames.message.document.body.innerHTML;"> 
              <input onClick="ClearReset()" type=reset name="Clear" value="重新填写"></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<!--#include file="foot.asp"-->
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
function checkkk()
{
     if(checkspace(document.myform.bookname.value)) {
	document.myform.bookname.focus();
    alert("请输入商品名称！");
	return false;
  }
     if(checkspace(document.myform.shichangjia.value)) {
	document.myform.shichangjia.focus();
    alert("请输入市场价格！");
	return false;
  }
     if(checkspace(document.myform.huiyuanjia.value)) {
	document.myform.huiyuanjia.focus();
    alert("请输入会员价格！");
	return false;
  }
     if(checkspace(document.myform.vipjia.value)) {
	document.myform.vipjia.focus();
    alert("请输入VIP会员价格！");
	return false;
  }
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
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
