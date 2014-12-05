
<%
'option explicit
'dim conn,connstr,startime,db,rs,rs_s,rs_s1
startime=timer()
db="bbs/data/cnhww.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
%>
