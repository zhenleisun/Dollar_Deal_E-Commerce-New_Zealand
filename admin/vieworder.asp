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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>
<%dim dingdan,username
dingdan=request.QueryString("dan")
if not isnumeric(dingdan) then 
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
username=request.QueryString("username")
if InStr(username,"'")>0 then
response.write"<script>alert(""�Ƿ�����!"");location.href=""../index.asp"";</script>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")

rs.open "select products.bookid,products.bookname,products.shichangjia,products.huiyuanjia,orders.actiondate,orders.shousex,orders.userzhenshiname,products.shjiaid,orders.shouhuoname,orders.dingdan,orders.youbian,orders.chima,orders.yanse,orders.liuyan,orders.zhifufangshi,orders.songhuofangshi,orders.zhuangtai,orders.zonger,orders.useremail,orders.usertel,orders.shouhuodizhi,orders.bookcount,orders.danjia,orders.fapiao,orders.feiyong from products inner join orders on products.bookid=orders.bookid where orders.username='"&username&"' and dingdan='"&dingdan&"' ",conn,1,1

if rs.eof and rs.bof then
response.write "<p align=center><font color=red>�˶���������Ʒ�ѱ�����Աɾ�����޷�������ȷ���㡣<br>��������ȡ�������ֶ�ɾ���˶�����</font></p>"
response.write "<input type=button name=Submit3 value=ɾ������ onClick=""location.href='saveorder.asp?action=del&dan="&dingdan&"&username="&username&"'""> </center>"
response.End
end if
shjiaid=rs("shjiaid")
%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="4" align="center" background="../images/admin_bg_1.gif"><b><font color="#ffffff">�޸���Ʒ����</font></b></td>
</tr>
                          <tr bgcolor="#fbf4f4"> 
                            <td colspan="2" align="center"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="89%" align="center">
������Ϊ��<%=dingdan%> ����ϸ�������£�
                                    </td>
                                    <td width="11%" align="center">
									<input type="button" name="Submit4" value="�� ӡ" onClick="javascript:window.print()">
                                    </td>
                                  </tr>
                                </table>
                            </td>
                          </tr>
                          <tr> 
                            <td width="20%" bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>����״̬��</td>
                            <td width="80%" bgcolor="#fbf4f4"> 
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                  <form name="form1" method="post" action="saveorder.asp?dan=<%=dingdan%>&action=save&username=<%=username%>">
                                    <td > 
                                      <%zhuang()%><br>
									  <input type="submit" name="Submit" value="�޸Ķ���״̬">
									  <font color="#FF0000">����ֻ��ǰ̨�û���ѡ��&quot;�û��Ѿ�������&quot;�󣬹���Ա�ſ��Խ��к�ߵĲ�����</font></td>
                                  </form>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          <tr> 
                            <td width="13%" bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>��Ʒ�б�</td>
                            <td bgcolor="#fbf4f4"> 
                              <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
        <tr> 
          <td bgcolor="#fbf4f4" align="center">��Ʒ����</td>
          <td bgcolor="#fbf4f4" align="center">��������</td>
          <td width="9%" align="center" bgcolor="#fbf4f4" >����</td>
          <td width="10%" align="center" bgcolor="#fbf4f4">��ɫ</td>
          <td bgcolor="#fbf4f4" align="center">�� ��</td>
          <td bgcolor="#fbf4f4" align="center">���С��</td>
        </tr>
        <%zongji=0
								do while not rs.eof%>
        <tr> 
          <td bgcolor="#fbf4f4" style='PADDING-LEFT: 5px'><a href=../products.asp?id=<%=rs("bookid")%> ><%=trim(rs("bookname"))%></a></td>
          <td bgcolor="#fbf4f4"> <div align="center"><%=rs("bookcount")%></div></td>
          <td align="center" bgcolor="#fbf4f4"><%=trim(rs("chima"))%></td>
          <td align="center" bgcolor="#fbf4f4" ><%=trim(rs("yanse"))%></td>
          <td bgcolor="#fbf4f4"> <div align="center"><%=rs("danjia")&"Ԫ"%></div></td>
          <td bgcolor="#fbf4f4"> <div align="center"><%=rs("zonger")&"Ԫ"%></div></td>
        </tr>
        <%zongji=rs("zonger")+zongji
								feiyong=rs("feiyong")
		rs.movenext
		loop
		rs.movefirst%>
        <tr> 
          <td colspan="6" bgcolor="#fbf4f4"> <div align="right">�����ܶ<%=zongji%>Ԫ�����ã�<%=feiyong%>Ԫ�������ƣ�<%=zongji+feiyong%>Ԫ 
              &nbsp;&nbsp;&nbsp;&nbsp;</div></td>
        </tr>
      </table>
                            </td>
                          </tr>
                          <%set snsn=server.CreateObject("adodb.recordset")
	snsn.open "select * from ordersaward where username='"&username&"' and dingdan='"&dingdan&"'",conn,1,1
	if snsn.recordcount>0 then%>
                          <tr bgcolor="#FFFFFF"> 
                            <td width="13%" style='PADDING-LEFT: 10px' bgcolor="#fbf4f4">��Ʒ�б�</td>
                            <td bgcolor="#fbf4f4"> 
                              <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" >
                                <tr> 
                                <td align="center">��Ʒ����</td>
                                <td align="center">���û���</td>
                                </tr>
	<%
	while not snsn.eof%>
                                <tr> 
                                  <td bgcolor="#fbf4f4"> 
                                    <div align="center">
                                      <%
	set snsn1=server.CreateObject("adodb.recordset")
	snsn1.open "select * from jiangpin where bookid="&snsn("bookid"),conn,1,1
	if snsn1.recordcount=1 then
	response.write snsn1("bookname")
	end if
	rs_sh.close
	set rs_sh=nothing
	snsn1.close
	set snsn1=nothing%>
                                    </div></td>
                                  <td align="center" bgcolor="#fbf4f4"><%=snsn("jifen")%></td>
                                </tr>
	<%
	snsn.movenext
	wend%>
                              </table>
                            </td>
                          </tr>
                          <%end if
snsn.close
set snsn=nothing%>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�ջ���������</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("userzhenshiname"))%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�ջ���ַ��</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("shouhuodizhi"))%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�� �ࣺ</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("youbian"))%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>��ϵ�绰��</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("usertel"))%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�����ʼ���</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("useremail"))%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�ͻ���ʽ��</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'> 
                              <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("songhuofangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.Close
    set rs2=nothing%>
                            </td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>֧����ʽ��</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'> 
                              <%set rs2=server.CreateObject("adodb.recordset")
    rs2.Open "select * from deliver where songid="&rs("zhifufangshi"),conn,1,1
    response.Write trim(rs2("subject"))
    rs2.close
    set rs2=nothing%>
                            </td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�Ƿ񿪷�Ʊ��</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%if rs("fapiao")=1 then %><font color=red>��Ҫ��Ʊ</font><%else%>����Ҫ��Ʊ<%end if%></td>
                          </tr>
                          <tr> 
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�û����ԣ�</td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=trim(rs("liuyan"))%></td>
                          </tr>
                          <tr> 
                            <td height="20" bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>�µ����ڣ�</td>
                            <td height="20" bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'><%=rs("actiondate")%></td>
                          </tr>
                          <tr>
						  <td height="30" bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'></td>
                            <td bgcolor="#fbf4f4" style='PADDING-LEFT: 10px'>
							<input type="button" name="Submit3" value="ɾ������" onClick="if(confirm('��ȷ��Ҫɾ����?')) location.href='saveorder.asp?action=del&dan=<%=dingdan%>&username=<%=username%>';else return;">
							&nbsp;
							<input type="button" name="Submit2" value="�رմ���" onClick=javascript:window.close()>
                            </td>
                          </tr>
                        </table>
<%sub zhuang()
select case rs("zhuangtai")
case "1"%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>δ���κδ��� �� 
<%if rs("zhifufangshi")<>4 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="checkbox" DISABLED>�û��Ѿ������� �� 
<input type="checkbox" name="zhuangtai" value="3">�������Ѿ��յ��� �� 
<input type="checkbox" name="checkbox3" value="checkbox" DISABLED>�������Ѿ����� �� 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>�û��Ѿ��յ��� 
<%else%>
<input name="zhuangtai" type="checkbox" id="zhuangtai" value="3">������ȷ�� �� 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>��������� 
<%end if%>
<%case "2"%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>δ���κδ��� �� 
<input name="checkbox" type="checkbox" id="checkbox" value="adf" checked DISABLED>�û��Ѿ������� �� 
<input type="checkbox" name="zhuangtai" value="3" >�������Ѿ��յ��� �� 
<input type="checkbox" name="checkbox" value="checkbox" DISABLED>�������Ѿ����� �� 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>�û��Ѿ��յ���
<%case "3"%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>δ���κδ��� �� 
<%if rs("zhifufangshi")<>4 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="checkbox" checked DISABLED>�û��Ѿ������� �� 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>�������Ѿ��յ��� �� 
<input type="checkbox" name="zhuangtai" value="4" >�������Ѿ����� �� 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>�û��Ѿ��յ���
<%else%>
<input name="checkbox" type="checkbox" DISABLED value="checkbox" checked>������ȷ�� �� 
<input type="checkbox" name="zhuangtai" id="checkbox" value="5">
��������� 
<%end if%>
<%case "4"%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>δ���κδ��� �� 
<input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>�û��Ѿ������� �� 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>�������Ѿ��յ��� �� 
<input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>�������Ѿ����� �� 
<input type="checkbox" name="checkbox" value="checkbox" DISABLED>�û��Ѿ��յ���
<%case "5"%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>δ���κδ��� �� 
<%if rs("zhifufangshi")<>4 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>�û��Ѿ������� �� 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>�������Ѿ��յ��� �� 
<input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>�������Ѿ����� �� 
<input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>�û��Ѿ��յ���
<%else%>
<input name="zhuangtai" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>��������� �� 
<input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>��������� 
<%end if%>
<%
end select
end sub%>
