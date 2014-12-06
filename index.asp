<HTML>
<HEAD>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<TITLE><%=webname%>--Home</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<meta name="description" content="">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFFFFF  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="text-align:center;">
<!--#include file="include/head.asp"-->
<div class="main">
  <table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="236" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><a href="index.asp"><img src="images/ds_08.jpg" width="236" height="48"></a></td>
        </tr>
        <tr>
          <td height="803" align="left" valign="top" background="images/ds_11.jpg"><table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
			  <tbody>
				<tr>
					<td  height="4" align="left"></td>
				</tr>
				<tr>
				  <td>
					<%
					set rs=server.CreateObject("adodb.recordset")
					rs.open "select  * from bsort order by anclassidorder",conn,1,1
					if rs.recordcount=0 then
						response.write "<br>There is no commodity classification"
					else
						while not rs.eof
					%>
					  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
						<tr>
						  <td colspan="3" height="30" align="left"><font color="#FF6600">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="class.asp?lx=big&amp;anid=<%=rs("anclassid")%>"><font color="#FF4800" size="3"><%=rs("anclass")%></font></a></font></td>
						</tr>
						<%
					set rs_s=server.CreateObject("adodb.recordset")
					rs_s.open "select * from ssort where anclassid=" & rs("anclassid") & " order by nclassidorder",conn,1,1
					if rs_s.recordcount=0 then 
			%>
						<tr>
						  <td colspan="3" height="22">No small class</td>
						</tr>
						<%
					else
						i=0
					while not rs_s.eof
			%>
						<tr>
						  <td  width="48%" height="28" align="right" bordercolor="#1941A5"><a href="class.asp?lx=small&amp;anid=<%=rs("anclassid")%>&amp;nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a></td>
						  <td  width="4%" align="center"><font color="ff6600"><b>|</b></font></td>
						  <%rs_s.movenext
			if rs_s.eof then 
			response.write "<td align=center>&nbsp;</td></tr>"
			else
			%>
						  <td align="left" width="48%" ><a href="class.asp?lx=small&amp;anid=<%=rs("anclassid")%>&amp;nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a></td>
						</tr>
						<%
			rs_s.movenext
			end if%>
						<%
					wend
					end if
			%>
					  </table>
					<%
					rs_s.close
					set rs_s=nothing
				rs.movenext
				wend
			end if
			rs.close
			set rs=nothing
			%></td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </tbody>
			</table></td>
        </tr>
        <tr>
          <td><img src="images/ds_14.jpg" width="236" height="35"></td>
        </tr>
      </table></td>
      <td width="725" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="434" align="left" valign="top" background="images/ds_09.jpg"><div style="padding:13px 0 0 7px; text-align:left;"><%
		set rs=server.CreateObject("adodb.recordset")
		rs.Open "select * from advertisement",conn,1,1
		tupian1=trim(rs("pic1"))
		tupian1url=trim(rs("pic1_lnk"))
		piaos1=trim(rs("tit1"))
		
		tupian2=trim(rs("pic2"))
		tupian2url=trim(rs("pic2_lnk"))
		piaos2=trim(rs("tit2"))
		
		tupian3=trim(rs("pic3"))
		tupian3url=trim(rs("pic3_lnk"))
		piaos3=trim(rs("tit3"))
		
		tupian4=trim(rs("pic4"))
		tupian4url=trim(rs("pic4_lnk"))
		piaos4=trim(rs("tit4"))
		 
		tupian5=trim(rs("pic5"))
		tupian5url=trim(rs("pic5_lnk"))
		 
		
		rs.Close
		set rs=nothing
		%>
		<script type="text/javascript">
		imgUrl1="<%=tupian1%>";
		imgtext1=""
		imgLink1=escape("<%=tupian1url%>");
		imgUrl2="<%=tupian2%>";
		imgtext2=""
		imgLink2=escape("<%=tupian2url%>");
		imgUrl3="<%=tupian3%>";
		imgtext3=""
		imgLink3=escape("<%=tupian3url%>");
		imgUrl4="<%=tupian4%>";
		imgtext4=""
		imgLink4=escape("<%=tupian4url%>");
		imgUrl5="<%=tupian5%>";
		imgtext5=""
		imgLink5=escape("<%=tupian5url%>");
		
		 var focus_width=704
		 var focus_height=415
		 var text_height=0
		 var swf_height = focus_height+text_height
		 
		 var pics=imgUrl1+"|"+imgUrl2+"|"+imgUrl3+"|"+imgUrl4+"|"+imgUrl5
		 var links=imgLink1+"|"+imgLink2+"|"+imgLink3+"|"+imgLink4+"|"+imgLink5
		 var texts=imgtext1+"|"+imgtext2+"|"+imgtext3+"|"+imgtext4+"|"+imgtext5
		 
		 document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
		 document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/focus2.swf"><param name="quality" value="high"><param name="bgcolor" value="#F0F0F0">');
		 document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
		 document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
		 document.write('<embed src="images/focus2.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="1002" height="423" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"></embed>');
		 document.write('</object>');
		</script></div></td>
        </tr>
        <tr>
          <td height="452" align="center" valign="top" background="images/ds_13.jpg"><!--#include file="include/newproduct.asp"--></td>
        </tr>
      </table></td>
    </tr>
  </table>
	<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr>
		<td align="center" valign="top" bgcolor="#FFFFFF"><!--#include file="include/cjcount.asp"--></td>
	  </tr>
  </table>
  <table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr>
		<td align="center" bgcolor="#FFFFFF"><!--#include file="include/special.asp"--></td>
	  </tr>
  </table>
  <table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr>
		<td height="40" align="center">&nbsp;</td>
	  </tr>
  </table>
</div>
<!--#include file="include/foot.asp"-->
</BODY>
</HTML>
