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
        excelfile.WriteLine("��ƷID,��Ʒ����,����Ʒ��,��Ʒ��λ,��Ʒ����,��Ʒ���,��ƷͼƬ,�г���,��Ա��,VIP��,�Ƿ�����,�Ƿ��ؼ�,�Ƿ���Ʒ,�������,�������,��Ʒ���")
    else
        excelfile.WriteLine("��ƷID,��Ʒ����,����Ʒ��,��Ʒ��λ,��Ʒ����,��Ʒ���,��ƷͼƬ,�г���,��Ա��,VIP��,�Ƿ�����,�Ƿ��ؼ�,�Ƿ���Ʒ,�������,�������,��Ʒ���")
    end if
    while not rs.eof
	if rs(10)=1 then 
	k="��"
	else k="��"
	end if 
		if rs(11)=1 then 
	t="��"
	else t="��"
	end if 
			if rs(12)=1 then 
	v="��"
	else v="��"
	end if
        excelfile.WriteLine(rs(0)&","&rs(1)&","&rs(2)&","&rs(3)&","&rs(4)&","&rs(5)&","&rs(6)&","&rs(7)&"Ԫ,"&rs(8)&"Ԫ,"&rs(9)&"Ԫ,"&k&","&t&","&v&","&rs(13)&","&rs(14)&","&rs(15))
        rs.movenext
    wend

excelfile.close
set fs=nothing
%>
<script>
    opener.document.all.excelfile.innerHTML="<a href='excel/<%=sRand%>' target=_blank><%=sRand%></a>";
    window.close();
</script>