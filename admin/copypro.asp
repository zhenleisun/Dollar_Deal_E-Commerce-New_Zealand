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
<%

dim action,bookid
bookid=request.QueryString("id")

set rse=server.CreateObject("adodb.recordset")
rse.Open "select * from products where bookid="&bookid,conn,1,3

set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from products ",conn,1,3

rs.AddNew

rs("anclassid")=rse("anclassid")
rs("nclassid")=rse("nclassid")
rs("xclassid")=rse("xclassid") '子类
rs("bookname")=rse("bookname")


set rsq=server.CreateObject("adodb.recordset")
rsq.Open "select top 1  from products",conn,1,3
id=rs("bookid")
   bianhao1=date()
        bianhao1=replace(bianhao1, "-", "")
        bianhao1=replace(bianhao1, " ", "")
        bianhao1=replace(bianhao1, ":", "")
        bianhao=trim(request("bianhao"))
		rs("grade")=bianhao1&id


rs("pingpai")=rse("pingpai")
rs("isbn")=rse("isbn") 
rs("chima")=rse("chima")
rs("yanse")=rse("yanse")
rs("bookchuban")=rse("bookchuban")
rs("shichangjia")=rse("shichangjia")  
rs("huiyuanjia")=rse("huiyuanjia") 
rs("vipjia")=rse("vipjia") 
rs("dazhe")=rse("dazhe") 
rs("kucun")=rse("kucun")  
rs("bookpic")=rse("bookpic")
rs("zhuang")=rse("zhuang") 
rs("metad")=rse("metad")	
rs("metak")=rse("metak")
rs("bookcontent")=rse("bookcontent")
rs("yeshu")=rse("yeshu")

rs("bestbook")=rse("bestbook")
rs("newsbook")=rse("newsbook")
rs("tejiabook")=rse("tejiabook")
rs("chengjiaocount")=rse("chengjiaocount") 
rs("liulancount")=rse("liulancount")  
rs("adddate")=now()
rs("pingji")=rse("pingji")  
rs("pingjizong")=rse("pingjizong")
rs.Update

rs.Close
set rs=nothing
if request("linkaddress")="" then
response.redirect request.servervariables("http_referer")
else
response.redirect request("linkaddress")
end if
response.End

%>