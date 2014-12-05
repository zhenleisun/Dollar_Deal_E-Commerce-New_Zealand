<!--#include file="conn.asp"-->
<%
if request("bookname")="" or request("shichangjia")="" or request("huiyuanjia")=""  then
response.Write "对不起，添加失败，请用正确的方式添加商品！"
response.End
end if
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
dim bookdate,dazhe
dazhe=round(request("huiyuanjia")/request("shichangjia"),2)
dim action,bookid
bookid=request.QueryString("id")
action=request.QueryString("action")
bookname=trim(Request.Form("bookname"))
bookname1=trim(Request.Form("bookname1"))
shjianame=request.cookies("cnhww")("shjianame")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from products",conn,1,3
rs.AddNew
rs("mch")=trim(request("mch")) 
rs("pp")=trim(request("pp")) 
rs("isbn1")=trim(request("isbn1")) 
rs("jj")=trim(request("jj")) 
rs("jg")=trim(request("jg")) 
rs("zl")=trim(request("zl")) 
rs("tp")=trim(request("tp")) 
rs("nr")=trim(request("nr")) 
rs("bh")=trim(request("bh")) 
rs("anclassid")=int(request("anclassid")) 
rs("nclassid")=int(request("nclassid")) 
rs("bookname")=trim(request("bookname")) 
rs("grade")=trim(request("grade")) 
rs("pingpai")=trim(request("pingpai"))
rs("chima")=trim(request("chima"))
rs("yanse")=trim(request("yanse"))
rs("isbn")=trim(request("isbn")) 
rs("bookchuban")=trim(request("bookchuban"))
rs("shichangjia")=trim(request("shichangjia"))  
rs("huiyuanjia")=trim(request("huiyuanjia")) 
rs("vipjia")=trim(request("vipjia")) 
rs("dazhe")=dazhe  
rs("kucun")=trim(request("kucun"))  
rs("bookpic")=trim(request("bookpic")) 
rs("zhuang")=trim(request("zhuang")) 
rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  
rs("yeshu")=trim(request("yeshu")) 
if request("bestbook")=1 then  
rs("bestbook")=1
else
rs("bestbook")=0
end if
if request("newsbook")=1 then  
rs("newsbook")=1
else
rs("newsbook")=0
end if
if request("tejiabook")=1 then  
rs("tejiabook")=1
else
rs("tejiabook")=0
end if

rs("chengjiaocount")=0  
rs("liulancount")=0  
rs("adddate")=now()  
rs("pingji")=0  
rs("pingjizong")=0 
rs.Update
bookid=rs("bookid")
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('添加成功！');history.go(-2);</script>"
response.End




case "Batch"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from products",conn,1,3

bookname1=bookname&","&bookname1
bookname2=split(bookname1,",")

for i=0 to ubound(bookname2)

if trim(bookname2(i))<>"" then

rs.AddNew
rs("mch")=trim(request("mch")) 
rs("pp")=trim(request("pp")) 
rs("isbn1")=trim(request("isbn1")) 
rs("jj")=trim(request("jj")) 
rs("jg")=trim(request("jg")) 
rs("zl")=trim(request("zl")) 
rs("tp")=trim(request("tp")) 
rs("nr")=trim(request("nr")) 
rs("bh")=trim(request("bh")) 
rs("bookname")=trim(bookname2(i))
rs("anclassid")=int(request("anclassid")) 
rs("nclassid")=int(request("nclassid")) 


if trim(request("grade"))="" then 
set rsq=server.CreateObject("adodb.recordset")
rsq.Open "select top 1  from products",conn,1,3
id=rs("bookid")
   bianhao1=date()
        bianhao1=replace(bianhao1, "-", "")
        bianhao1=replace(bianhao1, " ", "")
        bianhao1=replace(bianhao1, ":", "")
        bianhao=trim(request("bianhao"))
		rs("grade")=bianhao1&id
		else
rs("grade")=trim(request("grade")) 
end if 
rs("pingpai")=trim(request("pingpai"))
rs("isbn")=trim(request("isbn")) 

rs("bookchuban")=trim(request("bookchuban"))
rs("shichangjia")=trim(request("shichangjia"))  
rs("huiyuanjia")=trim(request("huiyuanjia")) 
rs("vipjia")=trim(request("vipjia")) 
rs("dazhe")=dazhe  
rs("chima")=trim(request("chima"))
rs("yanse")=trim(request("yanse"))
rs("kucun")=trim(request("kucun"))  
rs("bookpic")=trim(request("bookpic")) 
rs("zhuang")=trim(request("zhuang")) 

rs("bookcontent")=htmlencode2(trim(request("bookcontent")))  
rs("yeshu")=trim(request("yeshu")) 
if request("bestbook")=1 then  
rs("bestbook")=1
else
rs("bestbook")=0
end if
if request("newsbook")=1 then  
rs("newsbook")=1
else
rs("newsbook")=0
end if
if request("tejiabook")=1 then  
rs("tejiabook")=1
else
rs("tejiabook")=0
end if

rs("chengjiaocount")=0  
rs("liulancount")=0  
rs("adddate")=now()  
rs("pingji")=0  
rs("pingjizong")=0 
rs.Update

end if 
Next
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('添加成功！');history.go(-2);</script>"
response.End



case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from products where bookid="&bookid,conn,1,3
rs("mch")=trim(request("mch")) 
rs("pp")=trim(request("pp")) 
rs("isbn1")=trim(request("isbn1")) 
rs("jj")=trim(request("jj")) 
rs("jg")=trim(request("jg")) 
rs("zl")=trim(request("zl")) 
rs("tp")=trim(request("tp")) 
rs("nr")=trim(request("nr")) 
rs("bh")=trim(request("bh")) 
rs("anclassid")=int(request("anclassid")) 
rs("nclassid")=int(request("nclassid")) 
rs("bookname")=trim(request("bookname")) 
rs("pingpai")=trim(request("pingpai")) 
rs("grade")=trim(request("grade")) 
rs("isbn")=trim(request("isbn"))
rs("bookchuban")=trim(request("bookchuban")) 
rs("shichangjia")=trim(request("shichangjia"))  
rs("huiyuanjia")=trim(request("huiyuanjia"))  
rs("vipjia")=trim(request("vipjia")) 
rs("dazhe")=dazhe  
rs("chima")=trim(request("chima"))
rs("yanse")=trim(request("yanse"))
rs("kucun")=trim(request("kucun")) 
rs("bookpic")=trim(request("bookpic"))  
rs("zhuang")=trim(request("zhuang")) 
rs("bookcontent")=trim(request("bookcontent"))  
rs("yeshu")=trim(request("yeshu"))  
if request("bestbook")=1 then  
rs("bestbook")=1
else
rs("bestbook")=0
end if
if request("newsbook")=1 then  
rs("newsbook")=1
else
rs("newsbook")=0
end if
if request("tejiabook")=1 then  
rs("tejiabook")=1
else
rs("tejiabook")=0
end if
rs.Update
rs.Close
set rs=nothing
response.Write "<script language=javascript>alert('修改成功！');history.go(-2);</script>"
response.End
end select
%>
