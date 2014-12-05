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
<%dim action,html
html=request("html")
action=request.QueryString("action")
Content=Request.Form("Ccontent")
 
set rs=server.CreateObject("adodb.recordset")
select case action
case "huikuanfangshi"
rs.Open "select huikuanfangshi from webinfo",conn,1,3
rs("huikuanfangshi")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('付款方式修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "jiaoyitiaokuan"
rs.Open "select jiaoyitiaokuan from webinfo",conn,1,3
rs("jiaoyitiaokuan")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('交易条款修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "yunshushuoming"
rs.Open "select yunshushuoming from webinfo",conn,1,3
rs("yunshushuoming")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('运输说明修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "gouwuliucheng"
rs.Open "select gouwuliucheng from webinfo",conn,1,3
rs("gouwuliucheng")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('购物流程修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "changjianwenti"
rs.Open "select changjianwenti from webinfo",conn,1,3
rs("changjianwenti")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('常见问题修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "baomi"
rs.Open "select baomi from webinfo",conn,1,3
rs("baomi")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('保密和安全修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "shouhoufuwu"
rs.Open "select shouhoufuwu from webinfo",conn,1,3
rs("shouhoufuwu")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('商品销售和售后服务修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "shiyongfalv"
rs.Open "select shiyongfalv from webinfo",conn,1,3
rs("shiyongfalv")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('适用法律和版权声明修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "regtiaoyue"
rs.Open "select regtiaoyue from webinfo",conn,1,3
rs("regtiaoyue")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('注册条约修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "regtiaoyue2"
rs.Open "select regtiaoyue2 from webinfo",conn,1,3
rs("regtiaoyue2")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('注册条约修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "songhuofeiyong"
rs.Open "select songhuofeiyong from webinfo",conn,1,3
rs("songhuofeiyong")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('送货方式及费率修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "gongzuoshijian"
rs.Open "select gongzuoshijian from webinfo",conn,1,3
rs("gongzuoshijian")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('我们的工作时间修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "jifen"
rs.Open "select jifen from webinfo",conn,1,3
rs("jifen")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('积分奖励修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "vip"
rs.Open "select vip from webinfo",conn,1,3
rs("vip")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('VIP特惠修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "about"
rs.Open "select about from webinfo",conn,1,3
rs("about")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('关于本站修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "lxwm"
rs.Open "select lxwm from webinfo",conn,1,3
rs("lxwm")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('联系我们修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "jtcg"
rs.Open "select jtcg from webinfo",conn,1,3
rs("jtcg")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('集团采购修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "baozhuangbiaozhun"
rs.Open "select baozhuangbiaozhun from webinfo",conn,1,3
rs("baozhuangbiaozhun")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('个人包装标准修改成功');window.location.href='edithelp.asp';</script>"
response.End
case "baozhuangshuoming"
rs.Open "select baozhuangshuoming from webinfo",conn,1,3
rs("baozhuangshuoming")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('包装说明修改成功');window.location.href='edithelp.asp';</script>"
response.End
end select
%>
