<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
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
response.Write "<script language=javascript>alert('���ʽ�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "jiaoyitiaokuan"
rs.Open "select jiaoyitiaokuan from webinfo",conn,1,3
rs("jiaoyitiaokuan")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���������޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "yunshushuoming"
rs.Open "select yunshushuoming from webinfo",conn,1,3
rs("yunshushuoming")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('����˵���޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "gouwuliucheng"
rs.Open "select gouwuliucheng from webinfo",conn,1,3
rs("gouwuliucheng")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���������޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "changjianwenti"
rs.Open "select changjianwenti from webinfo",conn,1,3
rs("changjianwenti")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���������޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "baomi"
rs.Open "select baomi from webinfo",conn,1,3
rs("baomi")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���ܺͰ�ȫ�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "shouhoufuwu"
rs.Open "select shouhoufuwu from webinfo",conn,1,3
rs("shouhoufuwu")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('��Ʒ���ۺ��ۺ�����޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "shiyongfalv"
rs.Open "select shiyongfalv from webinfo",conn,1,3
rs("shiyongfalv")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���÷��ɺͰ�Ȩ�����޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "regtiaoyue"
rs.Open "select regtiaoyue from webinfo",conn,1,3
rs("regtiaoyue")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('ע����Լ�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "regtiaoyue2"
rs.Open "select regtiaoyue2 from webinfo",conn,1,3
rs("regtiaoyue2")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('ע����Լ�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "songhuofeiyong"
rs.Open "select songhuofeiyong from webinfo",conn,1,3
rs("songhuofeiyong")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('�ͻ���ʽ�������޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "gongzuoshijian"
rs.Open "select gongzuoshijian from webinfo",conn,1,3
rs("gongzuoshijian")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���ǵĹ���ʱ���޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "jifen"
rs.Open "select jifen from webinfo",conn,1,3
rs("jifen")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���ֽ����޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "vip"
rs.Open "select vip from webinfo",conn,1,3
rs("vip")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('VIP�ػ��޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "about"
rs.Open "select about from webinfo",conn,1,3
rs("about")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���ڱ�վ�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "lxwm"
rs.Open "select lxwm from webinfo",conn,1,3
rs("lxwm")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('��ϵ�����޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "jtcg"
rs.Open "select jtcg from webinfo",conn,1,3
rs("jtcg")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���Ųɹ��޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "baozhuangbiaozhun"
rs.Open "select baozhuangbiaozhun from webinfo",conn,1,3
rs("baozhuangbiaozhun")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('���˰�װ��׼�޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
case "baozhuangshuoming"
rs.Open "select baozhuangshuoming from webinfo",conn,1,3
rs("baozhuangshuoming")=content
rs.Update
session("content")=""
response.Write "<script language=javascript>alert('��װ˵���޸ĳɹ�');window.location.href='edithelp.asp';</script>"
response.End
end select
%>
