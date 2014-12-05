<%dim ad,adurl,ad1,ad2,ad3,ad4,ad5,ad6,ad7,ad8,ad1url,ad2url,ad3url,ad4url,ad5url,ad6url,ad7url,ad8url
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from advertisement",conn,1,1
ad=trim(rs("ad"))
adurl=trim(rs("adurl"))
ad1=trim(rs("ad1"))
ad1url=trim(rs("ad1url"))
ad2=trim(rs("ad2"))
ad2url=trim(rs("ad2url"))
ad3=trim(rs("ad3"))
ad3url=trim(rs("ad3url"))
ad4=trim(rs("ad4"))
ad4url=trim(rs("ad4url"))
ad5=trim(rs("ad5"))
ad5url=trim(rs("ad5url"))
ad6=trim(rs("ad6"))
ad6url=trim(rs("ad6url"))
ad7=trim(rs("ad7"))
ad7url=trim(rs("ad7url"))
ad8=trim(rs("ad8"))
ad8url=trim(rs("ad8url"))
dh=trim(rs("dh"))
dhurl=trim(rs("dhurl"))
rs.close
set rs=nothing
%>
<%
set rs=server.CreateObject("adodb.recordset")
rs.Open "select webname,webemail,dizhi,youbian,ip,lockip,nocopy,dhxg,fahuo,daohang,webnow,gdong,xp,tjj,tejia,webmess,dianhua,copyright,gonggao,weblogo,weburl,webbanner,icp,musicurl,alipaytype,qq from webinfo",conn,1,1
webname=trim(rs("webname"))
webemail=trim(rs("webemail"))
dizhi=trim(rs("dizhi"))
youbian=trim(rs("youbian"))
daohang=trim(rs("daohang"))
dianhua=trim(rs("dianhua"))
xp=trim(rs("xp"))
tjj=trim(rs("tjj"))
tejia=trim(rs("tejia"))
copyright=trim(rs("copyright"))
weblogo=trim(rs("weblogo"))
dhxg=trim(rs("dhxg"))
fahuomess=trim(rs("fahuo"))
webbanner=trim(rs("webbanner"))
weburl=trim(rs("weburl"))
gonggao=trim(rs("gonggao"))
icp=trim(rs("icp"))
shopurl=trim(rs("musicurl"))
alipaytype=trim(rs("alipaytype"))
webbanner=trim(rs("webbanner"))
nocopy=trim(rs("nocopy"))
webnow=trim(rs("webnow"))
lockip=trim(rs("lockip"))
ip=trim(rs("ip"))
gdong=trim(rs("gdong"))
webmess=trim(rs("webmess"))
qq=trim(rs("qq"))



rs.Close
set rs=nothing
Rem 过滤HTML代码
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    fString=ChkBadWords(fString)
    HTMLEncode = fString
else
   HTMLEncode=fstring
end if
end function
Rem 过滤SQL非法字符
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function
Rem 判断数字是否整形
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function

if nocopy=0 then
%><script>document.oncontextmenu=new Function("event.returnValue=false")
document.onselectstart=new Function("event.returnValue=false")</script>
<%end if %>
<!--#include file="Lockvist.asp"-->