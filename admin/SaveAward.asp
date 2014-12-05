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
if request("bookname")="" or request("shichangjia")="" or request("point")="" then
response.Write "<script language='javascript'>alert('请用正确的方式添加或修改奖品！');history.go(-1);</script>"
response.End
end if
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
dim bookdate,dazhe
'dazhe=round(request("huiyuanjia")/request("shichangjia"),2)
'if request("bookdateyear")<>"" then
'bookdate=trim(request("bookdateyear"))&"年"&trim(request("bookdatemonth"))&"月"
'else
'bookdate=""
'end if
dim action,bookid
bookid=request.QueryString("id")
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from award",conn,1,3
rs.AddNew
rs("bookname")=trim(request("bookname")) '商品名称
rs("guige")=trim(request("guige")) '商品规格
rs("shichangjia")=trim(request("shichangjia"))  '市场价
rs("point")=trim(request("point"))  '积分
rs("bookpic")=trim(request("bookpic"))  '小图片地址
rs("bookpic2")=trim(request("bookpic2")) '大图片地址
rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  '简介
rs("adddate")=now() '加入日期
if request("xianshi")=1 then  '推荐
rs("xianshi")=1
else
rs("xianshi")=0
end if
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('添加成功！');window.location.href='AddAward.asp';</script>"
response.End
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from award where bookid="&bookid,conn,1,3
rs("bookname")=trim(request("bookname")) '商品名称
rs("guige")=trim(request("guige")) '商品规格
rs("shichangjia")=trim(request("shichangjia"))  '市场价
rs("point")=trim(request("point"))  '积分
rs("bookpic")=trim(request("bookpic"))  '小图片地址
rs("bookpic2")=trim(request("bookpic2")) '大图片地址
rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  '简介
'rs("adddate")="bookdate"    '加入日期
if request("xianshi")=1 then  '推荐
rs("xianshi")=1
else
rs("xianshi")=0
end if
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('修改成功！');window.location.href='ManageAward.asp';</script>"
response.End
end select
%>