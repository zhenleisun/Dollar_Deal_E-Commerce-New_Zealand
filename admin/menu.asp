<!-- ��ʹ�õ�����Ȥ���Ϲ���ϵͳ��Ѱ汾����Ѱ�������ܿ���ʹ�ã����߼������ǲ��߱��ģ������Ҫ���Ͽ��꣬�Ƽ�ʹ�ùٷ���ʽ�棬�ٷ���վ:http://www.cnhww.com �ͷ�QQ��81447932  --><%


REM ������Ŀ����
dim menu(8,10)
menu(0,0)="��������"
menu(0,1)="<a href=setinfo.asp target=right><font color="&colorx&">��վ����</a> | <a href=managead.asp target=right><font color="&colorx&">�������</a>"
menu(0,2)="<a href=managezipcode.asp target=right><font color="&colorx&">�ʱ�����</a> | <a href=edithelp.asp target=right><font color="&colorx&">������Ϣ</a>"
menu(0,3)="<a href=manageremit.asp?action=songhuo target=right><font color="&colorx&">�ͻ���ʽ</a> | <a href=manageremit.asp?action=zhifu target=right><font color="&colorx&">֧����ʽ</a>"
menu(0,4)="<a href=links.asp target=right><font color="&colorx&">�������</a> | <a href=onlinepay.asp target=right><font color="&colorx&">֧������</a>"
menu(0,5)="<a href=open.asp target=right><font color="&colorx&">��վ����</a> | <a href=loginlog.asp target=right><font color="&colorx&">��½��־</a>"
menu(0,6)="<a href=sitemap.asp target=right><font color="&colorx&">����Google SiteMaps</a> "
menu(0,7)="<a href=LockIp.asp target=right><font color="&colorx&">��������</a> | <a href=fahuo.asp target=right><font color="&colorx&">��������</a>"

menu(1,0)="��Ʒ����"
menu(1,1)="<a href=managebsort.asp target=right>�������</a> | <a href=managessort.asp target=right>С�����</a>"
menu(1,2)="<a href=addproduct.asp target=right><font color="&color&">�����Ʒ</a> | <a href=manageproduct.asp target=right><font color="&color&">�޸���Ʒ</a>"
menu(1,3)="<a href=ManageBrand.asp target=right><font color="&color&">Ʒ�ƹ���</a> | <a href=manageunit.asp?action=no target=right><font color="&color&">��λ����</a>"
menu(1,4)="<a href=Mmoveclass.asp target=right>���ת��</a> | <a href=editorder.asp?zhuangtai=0 target=right><font color="&color&">�������� </a>"
menu(1,5)="<a href=managereview.asp?action=all target=right>���۹���</a> | <a href=dmail.asp target=right><font color="&color&">��ַ���� </a>"
menu(1,6)="<a href=OutProduct.asp target=right>ȱ������</a> | <a href=searchword.asp target=right><font color="&colorx&">������Ϣ</a>" 

menu(1,7)="<a href=add_batch.asp target=right>��Ʒ��Ϣ������ӹ���</a>"
menu(1,8)="<a href=chkpro.asp target=right>��Ʒ��Ϣ�����޸Ĺ���</a>"
menu(1,9)="<a href=coverexcel.asp target=right>��Ʒ��Ϣ������������</a>"


menu(2,0)="��Ϣ����"
menu(2,1)="<a href=addnews.asp target=right><font color="&color&">�������</a> | <a href=editnews.asp target=right><font color="&color&">��������</a>"
menu(2,2)="<a href=addinforms.asp target=right><font color="&color&">�����Ѷ</a> | <a href=manageinforms.asp target=right>������Ѷ</a>"
menu(2,3)="<a href=affiche.asp target=right><font color="&color&">��������</a> | <a href=viewreturn.asp target=right>���Թ���</a>"
menu(2,4)="<a href=managevote.asp target=right><font color="&color&">ͶƱ����</a> | <a href=reportforms.asp target=right>����ͳ��</a>"



menu(3,0)="VIP����"
menu(3,1)="<a href=AddAward.asp target=right>��ӽ�Ʒ</a> | <a href=ManageAward.asp target=right>�鿴�޸�</a>"
menu(3,2)="<a href=PointToAward.asp target=right><font color="&color&">���ֶҽ�</a> | <a href=VipExplain.asp target=right><font color="&color&">VIP ˵��</a>"
menu(3,3)="<a href=VipActivity.asp target=right><font color="&color&">VIP �</a> | <a href=LuckVip.asp target=right><font color="&color&">���� VIP</a>"

menu(4,0)="�û�����"
menu(4,1)="<a href=manageuser.asp?action=all target=right>ע���û�����</a>"
menu(4,2)="<a href=manageuser.asp?action=niming target=right>�����û�����</a>"
menu(4,3)="<a href=manageadmin.asp target=right>��̨�û�����</a>"

menu(5,0)="ʡ�й���"
menu(5,1)="<a href=shengmanage.asp target=right><font color="&color&">ʡ����</a> | <a href=shimanage.asp target=right><font color="&color&">�й���</a>"

menu(6,0)="���ݴ���(Access)"
menu(6,1)="<a href=backdata.asp target=right><font color="&colorx&">���ݱ���</a> | <a href=compressdata.asp target=right><font color="&colorx&">����ѹ��</a>"
menu(6,2)="<a a href=../check.asp target=right><font color="&color&">�ռ�ռ��</a> | <a href=aspcheck.asp target=right><font color="&colorx&">ϵͳ����</a>"

menu(7,0)="���Ź���"
menu(7,1)="<a href=hand1.asp target=right>�ռ���</a> | <a href=hand2.asp target=right>׫д�¶���</a>"
menu(7,2)="<a href=hand3.asp target=right><font color="&color&">������</a> | <a href=hand4.asp target=right>ɾ������Ϣ</a>"

menu(8,0)="����ͳ��"
menu(8,1)="<a href=tj1.asp target=right>��������ͳ��</a>"
menu(8,2)="<a href=tj2.asp target=right><font color="&color&">ÿ�շ�����ϸ</a>"
%>
<title>����ҳ��</title>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR>
<style type=text/css>
<!--
BODY{
background-color:#0072BC;
font-family:Verdana, Tahoma, Arial;
FONT-SIZE: 9pt; 
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;
SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;
SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow;
}
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#111111; text-decoration:none; }
a:hover  { color:#6C70AA;text-decoration:underline; }
.sec_menu {BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid}
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#0066CC; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#596099; font-weight:bold; }
input,select,Textarea{
font-family:����,Verdana, Arial, Helvetica, sans-serif; font-size: 12px;}
}
-->
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left> 
<tr>
  <td valign=top> <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        
    <td height=42 valign=bottom> <img src="../images/title.gif" width=158 height=38> 
    </td>
      </tr>
    </table>
    
<table cellpadding=0 cellspacing=0 width=158 align=center>
  
      <tr> 
        <td style="display:"> <div class=sec_menu style="width:158"> 
            
            <table cellpadding=0 cellspacing=0 align=center width=158>
              <TR style="CURSOR: hand"><TD width="23"  height=21 class=menu_title valign="top">
<div align="right"><img src="../images/home.gif" width="16" height="16" align="absmiddle"><br></div></TD>
<TD width="55" class=menu_title><b><a target="_top" href="Index.Asp"> ������ҳ</a></b> </TD>
<TD width="22" class=menu_title valign="top"><div align="right"><img src="../images/exit.gif" width="16" height="16" align="absmiddle"><br></div></TD>
          <TD width="56" class=menu_title><b><a target="_top" href="logout.asp"> �˳�����</a></b> 
          </TD>
</TR></table></td>
    </table>
    <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=158>
              <tr>
                
            <td height=10></td>
              </tr>
            </table>
          </div>
<%
	dim j,i
	dim tmpmenu
	dim menuname
	dim menurl
for i=0 to ubound(menu,1)
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/title_bg_show.gif" id=menuTitle1 onClick="showsubmenu(<%=i%>)"> 
      <span><%=menu(i,0)%></span> </td>
  </tr>
  <tr>
    <td <%if i=0 or i=1 or i=2 or i=3 or i=4 or i=5 or i=6 or i=7 or i=8   or i=9 then%> style="display:" <%else%> style="display:none" <%end if%> id='submenu<%=i%>'>
    <!--<td style="display:none" id='submenu<%=i%>'>-->
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=158>
	<%
	for j=1 to ubound(menu,2)
	if isempty(menu(i,j)) then exit for
	%>
<tr><td height=20><img src=../images/bullet.gif border=0><%=menu(i,j)%></td></tr>
<%
	next
%>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=158>
<tr><td height=10></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%next%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        
        
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="../images/title_bg_show.gif" id=menuTitle1> 
      <span>ϵͳ��Ϣ</span> </td>
      </tr>
      <tr> 
        <td> <div class=sec_menu style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=158>
              <tr>
                <td height=20><!--<br>
				��Ȩ����:<BR>
              ����Ȥ���Ϲ���ϵͳ<br>
              ����������---ʱ�а�V14.0<BR>
              �ٷ���վ:<BR>
              ��www.cnhww.com<BR>
              <BR>--></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table><BR>