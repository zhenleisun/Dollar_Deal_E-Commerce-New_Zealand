<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--在线支付</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<!--#include file="include/head.asp"-->
<table width="90%" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%" valign="top"> <br><TABLE width=685 height="165" border=0 align="center" cellPadding=0 cellSpacing=0>
            <TBODY>
              <TR>
                <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_004.gif" width=4></TD>
                <TD background=images/new_line_008.gif height=4></TD>
                <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_005.gif" width=4></TD>
              </TR>
              <TR>
                <TD width="1" background=images/new_line_009.gif></TD>
                <TD 
          style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 3px; PADDING-TOP: 0px" 
          align=middle><img src="images/vendor_img.gif" width="672" height="150"></TD>
                <TD background=images/new_line_010.gif></TD>
              </TR>
              <TR>
                <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_006.gif" width=4></TD>
                <TD background=images/new_line_011.gif></TD>
                <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_007.gif" width=4></TD>
              </TR>
            </TBODY>
          </TABLE>
      <br><TABLE width=685 height="224" border=0 align="center" cellPadding=0 cellSpacing=0>
        <TBODY>
          <TR>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_004.gif" width=4></TD>
            <TD background=images/new_line_008.gif height=4></TD>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_005.gif" width=4></TD>
          </TR>
          <TR>
            <TD width="4" background=images/new_line_009.gif></TD>
            <TD 
          style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 3px; PADDING-TOP: 0px" 
          align=middle><table width="100%" height="224"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td height="224"><form method="post" action="checklist.asp" name="E_FORM">
                    <center>
                              <table width="444">
                                <tr> 
                                  <td align="right"> 姓　名</td>
                                  <td> <input name="dinghuorenxingming" type="text"  size="30"> 
                                  </td>
                                </tr>
                                <tr> 
                                  <td align="right"> 地　址</td>
                                  <td> <input name="shouhuorendizhi" type="text" size="50"> 
                                  </td>
                                </tr>
                                <tr> 
                                  <td align="right"> 电　话</td>
                                  <td> <div align="left"> 
                                      <input name="shouhuorendianhua" type="text" size="50">
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td align="right"> 邮　编</td>
                                  <td> <input name="shouhuorenyoubian" type="text"size="50"> 
                                  </td>
                                </tr>
                                <tr> 
                                  <td align="right">金　额</td>
                                  <td><input name="v_amount" type=text value="100.00" >
                                    元　 譬如：<font color="#FF0000">0.01</font></td>
                                </tr>
                                <tr>
                                  <td align="right">备　注</td>
                                  <td><textarea name="textarea" cols="48" rows="2"></textarea></td>
                                </tr>
                              </table>
                            </center>
                            <center>
                           
                        <input type=submit name=v_action value="提交" onClick="MM_validateForm('dinghuorenxingming','','R','shouhuorendizhi','','R','shouhuorendianhua','','R','shouhuorenyoubian','','RisNum');return document.MM_returnValue">
                      </p>
                    </center>
                </form></td>
              </tr>
            </table></TD>
            <TD width="5" background=images/new_line_010.gif></TD>
          </TR>
          <TR>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_006.gif" width=4></TD>
            <TD background=images/new_line_011.gif></TD>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_007.gif" width=4></TD>
          </TR>
        </TBODY>
      </TABLE><br>
      <TABLE width=672 height="157" border=0 align="center" cellPadding=0 cellSpacing=0>
        <TBODY>
          <TR>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_004.gif" width=4></TD>
            <TD background=images/new_line_008.gif height=4></TD>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_005.gif" width=4></TD>
          </TR>
          <TR>
            <TD width="1" background=images/new_line_009.gif></TD>
            <TD 
          style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 3px; PADDING-TOP: 0px" 
          align=middle><TABLE cellSpacing=0 cellPadding=0 width=720>
              <TBODY>
                <TR 
              style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 5px; PADDING-TOP: 5px" 
              vAlign=top height=185>
                  <TD>
                    <TABLE cellSpacing=0 cellPadding=0 width=349>
                      <TBODY>
                        <TR>
                          <TD colSpan=2><img height=64 
                        src="images/vendor_stit01.gif" 
                        width=349></TD>
                        </TR>
                        <TR height=89>
                          <TD class=goods style="PADDING-LEFT: 13px" 
                        width=225><img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 填写完整的资料<BR>
                              <img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 确认金额<BR>
                              <img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 提交</TD>
                          <TD width="124">&nbsp;</TD>
                        </TR>
                        <TR>
                          <TD 
                      background="images/dotline_h2.gif" 
                      colSpan=2 height=1></TD>
                        </TR>
                        <TR>
                          <TD colSpan=2 height=29><img hspace=2 
                        src="images/vendor_icon.gif" 
                        align=absMiddle> </TD>
                        </TR>
                      </TBODY>
                  </TABLE></TD>
                  <TD width=1 bgColor=#d7d7d7></TD>
                  <TD align=right>
                    <TABLE cellSpacing=0 cellPadding=0 width=349>
                      <TBODY>
                        <TR>
                          <TD colSpan=2><img height=64 
                        src="images/vendor_stit02.gif" 
                        width=349></TD>
                        </TR>
                        <TR height=89>
                          <TD class=goods style="PADDING-LEFT: 13px" 
                        width=249><img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 再次确认您的信息<BR>
                              <img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 若有错请点击会返回进行修改<BR>
                                    <img 
                        src="images/ico_rec_lightgreen2.gif" 
                        align=absMiddle vspace=6> 点击“支付”进入NPS支付系统进行付款</TD>
                          <TD width="100">&nbsp;</TD>
                        </TR>
                        <TR>
                          <TD 
                      background="images/dotline_h2.gif" 
                      colSpan=2 height=1></TD>
                        </TR>
                        <TR>
                          <TD colSpan=2 height=29><img hspace=2 
                        src="images/vendor_icon.gif" 
                        align=absMiddle> </TD>
                        </TR>
                      </TBODY>
                  </TABLE></TD>
                </TR>
              </TBODY>
            </TABLE></TD>
            <TD background=images/new_line_010.gif></TD>
          </TR>
          <TR>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_006.gif" width=4></TD>
            <TD background=images/new_line_011.gif></TD>
            <TD vAlign=top width=4 height=4><img height=4 
            src="images/new_line_007.gif" width=4></TD>
          </TR>
        </TBODY>
      </TABLE> <br>
      </td>
  </tr>
</table>
</td>
</tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>