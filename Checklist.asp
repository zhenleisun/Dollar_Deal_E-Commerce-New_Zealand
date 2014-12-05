<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="pay/md5nps.asp"-->
<html>
<head>
<title><%=webname%>--在线支付</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language='JavaScript'>
<!--
function conFirmv_amount()
{
	//document.netbank.v_amount.value=Math.round(document.netbank.v_amount.value);
	if ( document.netbank.v_amount.value <= 0 )
	{
	alert('金额必须大于0');
	return false;
	}
	//var stConFirm = '您的转帐金额为: ' + document.netbank.v_amount.value + ' 元!\n';
	//return confirm(stConFirm);
}
//-->
</script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<!--#include file="include/head.asp"-->
<table width="90%" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
<tr> 
  <td height="260" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table height="273"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="100%" valign="top"><br><TABLE width=685 height="165" border=0 align="center" cellPadding=0 cellSpacing=0>
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
        <br><TABLE width="685" height="200" border=0 align="center" cellPadding=0 cellSpacing=0>
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
          align=middle><table width="100%" height="200"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><%    exec="select * from webinfo"
      set rs=server.createobject("adodb.recordset")
      rs.open exec,conn,1,3
      npayid=rs("npayid")
      npaypin=rs("npaypin")
      npayurl=rs("npayurl")
     
	v_ordername=request.form("dinghuorenxingming")
	v_rcvaddr=request.form("shouhuorendizhi")
	v_rcvtel=request.form("shouhuorendianhua")
	v_rcvpost=request.form("shouhuorenyoubian")
	v_amount=trim(request.form("v_amount"))
	textarea=request.form("textarea")
%>
  
                      <br>
                      <table border="0" align=center>
                        <tr>
                          <td align="center">请再次确认您的付款信息：<br>
                              <br>
                              <table width="95%" border="0" bgcolor="#f0f0f0" class="tableBorder">
                                  <tr> 
                                    <td align="left" width="100">姓　名</td>
                                    <td width="300"><%=v_ordername%></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" width="100">地　址：</td>
                                    <td><%=v_rcvaddr%></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" width="100">电　话：</td>
                                    <td><%=v_rcvtel%></td>
                                  </tr>
                                  <tr> 
                                    <td align="left" width="100">邮　编：</td>
                                    <td><%=v_rcvpost%></td>
                                  </tr>
                                  <tr> 
                                    <td width="100" align="left"><font color="#FF6600">金　额:</font></td>
                                    <td><font color="#FF6600"><%=v_amount%></font></td>
                                  </tr>
                                  <tr>
                                    <td width="100" align="left">备　注：</td>
                                    <td><font color="#FF6600"><%=textarea%></font></td>
                                  </tr>
                                </table></td>
                        </tr>
                      </table>
                      <%


	curdate=now()
	m_orderid=year(curdate)&month(curdate)&day(curdate)&ipayid&hour(curdate)&minute(curdate)&second(curdate)


	m_id		=	npayid
	modate		=	date+time
	m_orderid	=	m_orderid
	m_oamount	=	v_amount
	m_ocurrency	=	"1"
	m_url		=	npayurl
	'm_url		=	"http://weburl/pay/NpsVirReceive.asp?do=result"
	m_language	=	"1"
	s_name		= v_ordername
	s_addr		= v_rcvaddr
	s_postcode	= v_rcvpost
	s_tel		= v_rcvtel			
	s_eml		=	"cnhww"
	r_name		= v_ordername
	r_addr		= v_rcvaddr
	r_postcode	= v_rcvpost
	r_tel		=	v_rcvtel
	r_eml		=	"cnhww@163.com"
	m_ocomment	= textarea	
	m_status	=	"0"
	key		=	npaypin
	
	OrderMessage =m_id&m_orderid&m_oamount&m_ocurrency&m_url&m_language&s_postcode&s_tel&s_eml&r_postcode&r_tel&r_eml&modate&key

	
	
	digest = Ucase(trim(md5(OrderMessage)))
	

	
%>

<div align=center>

<form method="post" action="https://payment.nps.cn/VirReceiveMerchantAction.do" target="_self">
      <input Type="hidden" Name="M_ID" value="<%=m_id%>">
	<input Type="hidden" Name="MOrderID" value="<%=m_orderid%>">
	<input Type="hidden" Name="MOAmount" value="<%=m_oamount%>">
	<input Type="hidden" Name="MOCurrency" value="<%=m_ocurrency%>">
	<input Type="hidden" name="M_URL" value="<%=m_url%>">
	<input Type="hidden" Name="M_Language" value="<%=m_language%>">
	<input Type="hidden" Name="S_Name" value="<%=s_name%>">
	<input Type="hidden" Name="S_Address" value="<%=s_addr%>">
    <input Type="hidden" Name="S_PostCode" value="<%=s_postcode%>">
	<input Type="hidden" Name="S_Telephone" value="<%=s_tel%>">
	<input Type="hidden" Name="S_Email" value="<%=s_eml%>">
	<input Type="hidden" Name="R_Name" value="<%=r_name%>"> 
	<input Type="hidden" Name="R_Address" value="<%=r_addr%>">
	<input Type="hidden" Name="R_PostCode" value="<%=r_postcode%>">
	<input Type="hidden" Name="R_Telephone" value="<%=r_tel%>">
	<input Type="hidden" Name="R_Email" value="<%=r_eml%>">
	<input Type="hidden" name="MOComment" value="<%=m_ocomment%>">
	<input Type="hidden" Name="MODate" value="<%=modate%>">
	<input Type="hidden" Name="State" value="<%=m_status%>">
	<input Type="hidden" Name="digestinfo" value="<%=digest%>">
	<input Type="submit" Name="submit" value="立即使用NPS在线支付"> 
</form>

                      <table border="0" width="100%">
                    </table></td>
                </tr>
              </table></TD>
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
                        align=absMiddle vspace=6> 点击“支付”进入网银支付系统进行付款</TD>
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
        </TABLE></td>
    </tr>
  </table></td>
</tr>
</table>
 
<!--#include file="include/foot.asp"-->
</body>
</html>