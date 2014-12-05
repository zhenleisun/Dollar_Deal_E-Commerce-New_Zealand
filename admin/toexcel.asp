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
set rs=server.createobject("adodb.recordset")
sql=request("sql")
rs.open sql,conn,1,1
curpath=Server.mappath("excel\")
Set fs = CreateObject("Scripting.FileSystemObject")
if not fs.FolderExists(curpath) then fs.CreateFolder(curpath)
Randomize
sRand=rnd
sRand=year(now)&month(now)&hour(now)&minute(now)&second(now)&sRand&".CSV"
Set excelfile = fs.CreateTextFile(curpath&"\"&sRand, True)

    if not rs.eof then
        excelfile.WriteLine("商品ID,商品名称,所属品牌,商品单位,商品积分,商品规格,商品图片,市场价,会员价,VIP价,是否热销,是否特价,是否新品,库存数量,添加日期,商品编号")
    else
        excelfile.WriteLine("商品ID,商品名称,所属品牌,商品单位,商品积分,商品规格,商品图片,市场价,会员价,VIP价,是否热销,是否特价,是否新品,库存数量,添加日期,商品编号")
    end if
    while not rs.eof
	if rs(10)=1 then 
	k="是"
	else k="否"
	end if 
		if rs(11)=1 then 
	t="是"
	else t="否"
	end if 
			if rs(12)=1 then 
	v="是"
	else v="否"
	end if
        excelfile.WriteLine(rs(0)&","&rs(1)&","&rs(2)&","&rs(3)&","&rs(4)&","&rs(5)&","&rs(6)&","&rs(7)&"元,"&rs(8)&"元,"&rs(9)&"元,"&k&","&t&","&v&","&rs(13)&","&rs(14)&","&rs(15))
        rs.movenext
    wend

excelfile.close
set fs=nothing
%>
<script>
    opener.document.all.excelfile.innerHTML="<a href='excel/<%=sRand%>' target=_blank><%=sRand%></a>";
    window.close();
</script>