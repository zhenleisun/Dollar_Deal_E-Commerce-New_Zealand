<table width="180" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/body/left03.gif" width="180" height="50" /></td>
  </tr>
  <tr>
    <td ><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr><%
		set rs=conn.execute("SELECT * from vote where IsChecked=1 ") 
		if rs.eof then
	  %>
                    ÔÝÎÞÍ¶Æ±
                    <%else%>
        <td><div align="center"><b> <%=rs("Title")%></b></div></td>
      </tr>
      <tr>
        <td><table border="0" width="100%" cellspacing="2" cellpadding="2">
                                                <form action="vote.asp" target="newwindow" method=post name=research>
                                                  <tr>
                                                    <td valign=top width="100%"><%
for i=1 to 8
	if rs("Select"&i)<>"" then
%>
                                                        <input style="border: 0" <%if i=1 then%>checked<%end if%> name=Options type=radio value=<%=i%>>
                                                        <%=i%>.<%=rs("Select"&i)%><br>
                                                        <%
	end if
next
%>                                                    </td>
                                                  </tr>
                                                  <tr>
                                                    <td width="100%" height=30 align=center><input type="image" border="0" name="submit" src="images/vote1.gif" width="59" height="19" style="cursor:hand"onClick="window.open('vote.asp?stype=view ','basket','menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbars=0,resizable=0,width=510,top=50,left=100,height=370')">
                                                       <a href=# onClick="window.open('vote.asp?stype=view ','basket','menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbars=0,resizable=0,width=510,top=50,left=100,height=370')"><img src="images/vote.gif" width="59" height="19" border="0" /> </a></td>
                                                  </tr>
                                                </form>
                                            </table>
                                                <%end if%>          </td>
      </tr>
    </table></td>
  </tr>
</table>
