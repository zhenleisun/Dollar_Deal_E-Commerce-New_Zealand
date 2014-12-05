<!-- 您使用的是网趣网上购物系统免费版本，免费版基本功能可以使用，诸多高级功能是不具备的，如果您要网上开店，推荐使用官方正式版，官方网站:http://www.cnhww.com 客服QQ：81447932  --><%


REM 管理栏目设置
dim menu(8,10)
menu(0,0)="常规设置"
menu(0,1)="<a href=setinfo.asp target=right><font color="&colorx&">网站设置</a> | <a href=managead.asp target=right><font color="&colorx&">广告设置</a>"
menu(0,2)="<a href=managezipcode.asp target=right><font color="&colorx&">邮编设置</a> | <a href=edithelp.asp target=right><font color="&colorx&">帮助信息</a>"
menu(0,3)="<a href=manageremit.asp?action=songhuo target=right><font color="&colorx&">送货方式</a> | <a href=manageremit.asp?action=zhifu target=right><font color="&colorx&">支付方式</a>"
menu(0,4)="<a href=links.asp target=right><font color="&colorx&">合作伙伴</a> | <a href=onlinepay.asp target=right><font color="&colorx&">支付参数</a>"
menu(0,5)="<a href=open.asp target=right><font color="&colorx&">网站开关</a> | <a href=loginlog.asp target=right><font color="&colorx&">登陆日志</a>"
menu(0,6)="<a href=sitemap.asp target=right><font color="&colorx&">生成Google SiteMaps</a> "
menu(0,7)="<a href=LockIp.asp target=right><font color="&colorx&">访问限制</a> | <a href=fahuo.asp target=right><font color="&colorx&">发货管理</a>"

menu(1,0)="商品管理"
menu(1,1)="<a href=managebsort.asp target=right>大类管理</a> | <a href=managessort.asp target=right>小类管理</a>"
menu(1,2)="<a href=addproduct.asp target=right><font color="&color&">添加商品</a> | <a href=manageproduct.asp target=right><font color="&color&">修改商品</a>"
menu(1,3)="<a href=ManageBrand.asp target=right><font color="&color&">品牌管理</a> | <a href=manageunit.asp?action=no target=right><font color="&color&">单位管理</a>"
menu(1,4)="<a href=Mmoveclass.asp target=right>类别转移</a> | <a href=editorder.asp?zhuangtai=0 target=right><font color="&color&">订单管理 </a>"
menu(1,5)="<a href=managereview.asp?action=all target=right>评论管理</a> | <a href=dmail.asp target=right><font color="&color&">邮址订阅 </a>"
menu(1,6)="<a href=OutProduct.asp target=right>缺货管理</a> | <a href=searchword.asp target=right><font color="&colorx&">搜索信息</a>" 

menu(1,7)="<a href=add_batch.asp target=right>商品信息批量添加功能</a>"
menu(1,8)="<a href=chkpro.asp target=right>商品信息批量修改功能</a>"
menu(1,9)="<a href=coverexcel.asp target=right>商品信息批量导出功能</a>"


menu(2,0)="信息管理"
menu(2,1)="<a href=addnews.asp target=right><font color="&color&">添加新闻</a> | <a href=editnews.asp target=right><font color="&color&">管理新闻</a>"
menu(2,2)="<a href=addinforms.asp target=right><font color="&color&">添加资讯</a> | <a href=manageinforms.asp target=right>管理资讯</a>"
menu(2,3)="<a href=affiche.asp target=right><font color="&color&">公告设置</a> | <a href=viewreturn.asp target=right>留言管理</a>"
menu(2,4)="<a href=managevote.asp target=right><font color="&color&">投票管理</a> | <a href=reportforms.asp target=right>销售统计</a>"



menu(3,0)="VIP管理"
menu(3,1)="<a href=AddAward.asp target=right>添加奖品</a> | <a href=ManageAward.asp target=right>查看修改</a>"
menu(3,2)="<a href=PointToAward.asp target=right><font color="&color&">积分兑奖</a> | <a href=VipExplain.asp target=right><font color="&color&">VIP 说明</a>"
menu(3,3)="<a href=VipActivity.asp target=right><font color="&color&">VIP 活动</a> | <a href=LuckVip.asp target=right><font color="&color&">幸运 VIP</a>"

menu(4,0)="用户管理"
menu(4,1)="<a href=manageuser.asp?action=all target=right>注册用户管理</a>"
menu(4,2)="<a href=manageuser.asp?action=niming target=right>匿名用户管理</a>"
menu(4,3)="<a href=manageadmin.asp target=right>后台用户管理</a>"

menu(5,0)="省市管理"
menu(5,1)="<a href=shengmanage.asp target=right><font color="&color&">省管理</a> | <a href=shimanage.asp target=right><font color="&color&">市管理</a>"

menu(6,0)="数据处理(Access)"
menu(6,1)="<a href=backdata.asp target=right><font color="&colorx&">数据备份</a> | <a href=compressdata.asp target=right><font color="&colorx&">数据压缩</a>"
menu(6,2)="<a a href=../check.asp target=right><font color="&color&">空间占用</a> | <a href=aspcheck.asp target=right><font color="&colorx&">系统环境</a>"

menu(7,0)="短信管理"
menu(7,1)="<a href=hand1.asp target=right>收件箱</a> | <a href=hand2.asp target=right>撰写新短信</a>"
menu(7,2)="<a href=hand3.asp target=right><font color="&color&">发件箱</a> | <a href=hand4.asp target=right>删除短消息</a>"

menu(8,0)="访问统计"
menu(8,1)="<a href=tj1.asp target=right>总体数据统计</a>"
menu(8,2)="<a href=tj2.asp target=right><font color="&color&">每日访问明细</a>"
%>
<title>管理页面</title>
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
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#111111; text-decoration:none; }
a:hover  { color:#6C70AA;text-decoration:underline; }
.sec_menu {BORDER-RIGHT: white 1px solid; BACKGROUND: #d6dff7; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid}
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#0066CC; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#596099; font-weight:bold; }
input,select,Textarea{
font-family:宋体,Verdana, Arial, Helvetica, sans-serif; font-size: 12px;}
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
<TD width="55" class=menu_title><b><a target="_top" href="Index.Asp"> 管理首页</a></b> </TD>
<TD width="22" class=menu_title valign="top"><div align="right"><img src="../images/exit.gif" width="16" height="16" align="absmiddle"><br></div></TD>
          <TD width="56" class=menu_title><b><a target="_top" href="logout.asp"> 退出管理</a></b> 
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
      <span>系统信息</span> </td>
      </tr>
      <tr> 
        <td> <div class=sec_menu style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=158>
              <tr>
                <td height=20><!--<br>
				版权所有:<BR>
              　网趣网上购物系统<br>
              　　　　　---时尚版V14.0<BR>
              官方网站:<BR>
              　www.cnhww.com<BR>
              <BR>--></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table><BR>