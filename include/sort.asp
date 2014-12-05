<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tbody>
    <tr>
      <td height="28"><div align="center"><img src="images/body/left02.gif" width="190" height="43" /></div></td>
    </tr>
    <tr>
      <td>
        <%
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select  * from bsort order by anclassidorder",conn,1,1
		if rs.recordcount=0 then
			response.write "<br>At present there is no commodity classification"
		else
			while not rs.eof
		%>
          <table width="100%" cellspacing="0" cellpadding="0" border="0">
            <tr>
              <td colspan="3" height="22" align="left"><font color="#FF6600">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="class.asp?lx=big&amp;anid=<%=rs("anclassid")%>"><font color="#FF4800"><strong><%=rs("anclass")%></strong></font></a></font></td>
            </tr>
            <%
	    set rs_s=server.CreateObject("adodb.recordset")
	    rs_s.open "select * from ssort where anclassid=" & rs("anclassid") & " order by nclassidorder",conn,1,1
	    if rs_s.recordcount=0 then 
%>
            <tr>
              <td colspan="3" height="22">No small classification</td>
            </tr>
            <%
	    else
	    	i=0
		while not rs_s.eof
%>
            <tr>
              <td  width="48%" height="20" align="right"><a href="class.asp?lx=small&amp;anid=<%=rs("anclassid")%>&amp;nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a></td>
              <td  width="4%" align="center"><font color="ff6600"><b>|</b></font></td>
              <%rs_s.movenext
if rs_s.eof then 
response.write "<td align=center>&nbsp;</td></tr>"
else
%>
              <td align="left" width="48%" height="22" ><a href="class.asp?lx=small&amp;anid=<%=rs("anclassid")%>&amp;nid=<%=rs_s("nclassid")%>"><%=rs_s("nclass")%></a></td>
            </tr>
            <%
rs_s.movenext
end if%>
            <%
		wend
	    end if
%>
          </table>
        <%
	    rs_s.close
	    set rs_s=nothing
	rs.movenext
	wend
end if
rs.close
set rs=nothing
%></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </tbody>
</table>
