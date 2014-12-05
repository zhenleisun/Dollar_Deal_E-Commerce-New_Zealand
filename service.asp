<TABLE cellSpacing=0 cellPadding=0 width=970 align=center border=0 bgColor=#f3f3f3>
  <TBODY>
    <TR>
      <TD class=b vAlign=top align=left width=970><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" >
        <tr>
          <td height="25" align="center" valign="middle" >ºÏ×÷»ï°é
              <%set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from links order by linkidorder",conn,1,1
	do while not rs.EOF
	response.Write " | <a href="&trim(rs("linkurl"))&" >"&trim(rs("linkname"))&"</a>"
	rs.MoveNext
	loop
	rs.Close
	set rs=nothing
	%>
          </td>
        </tr>
      </table></TD>
    </TR>
  </TBODY>
</TABLE>
