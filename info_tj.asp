<%
where=Request.ServerVariables("HTTP_REFERER")
IP = Request.ServerVariables("REMOTE_ADDR")
response.cookies("dmwh_user_ip")=ip

'更新访问统计基本数据
set rs=conn.Execute ("select * from count_total")
if rs("zzday")<>date() then
  conn.Execute("Update count_total set zzday=#"&date()&"#,yesterday=today,today=1,total=total+1")
else
  conn.Execute("Update count_total set today=today+1,total=total+1")
end if

'更新在线访客
set rs=server.createobject("adodb.recordset")
sql="select * from count_online where ip='"&IP&"' and datediff('h',time,now())<1"
rs.open sql,conn,1,3
if (rs.eof and rs.bof) then 
    rs.addnew
    rs("IP")=ip
    rs.update
end if
rs.close
set rs=nothing

'更新访问记录
set rs=server.createobject("adodb.recordset")
sql="select * from count_shop where  ip='"&IP&"' and date()=day and datediff('h',times,time())<2"
rs.open sql,conn,1,3
    if (rs.eof and rs.bof) then
    rs.addnew
    rs("IP")=IP
    rs("where")=lleft(where,200)
    rs.update
    end if
rs.close
set rs=nothing
%>