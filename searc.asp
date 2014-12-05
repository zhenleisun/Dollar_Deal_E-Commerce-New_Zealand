<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0" >
  <tbody>
    <tr> 
      <td align="left"><TABLE width="190" border=0 align="left" cellPadding=0 cellSpacing=0>
		  <TBODY>
		  <TR> 
              <TD align="center" bgcolor="#FFFFFF"><IMG height=50 src="images/body/searc.gif" width=190></TD>
		  </TR>
		  <TR> 
              <TD style="background:url(images/leftbg.jpg) repeat-y;"><TABLE align=center border=0 cellPadding=0 cellSpacing=0 width="100%">
				<FORM action=research.asp method=post name=form2>
				  <TBODY>
					<TR> 
					  <TD width="42%" align=right>Keyword£º </TD>
					  <TD width="58%" align="left"><INPUT class=wenbenkuang name=searchkey size=12 ;> </TD>
					</TR>
					<TR> 
					  <TD align=right>classify£º </TD>
					  <%set rs=server.CreateObject("adodb.recordset")
					  rs.open "select * from bsort order by anclassidorder",conn,1,1
					  %><TD align="left">
                        <select name="anclassid">
                              <option value="0">classify</option>
                              <%do while not rs.eof%>
                              <option value="<%=rs("anclassid")%>"> <%if len(trim(rs("anclass")))>5 then
								response.write left(trim(rs("anclass")),5)&"."
								else
								response.write trim(rs("anclass"))
								end if%></option>
                              <%rs.movenext
							  loop
							  rs.close
							  set rs=nothing%>
                       </select> </TD>
					</TR>
					<TR> 
					  <TD align=center height=38 colspan="2"><INPUT name="Submit" style="background:url(images/ioc_anniu.gif) no-repeat; cursor:pointer; color:#FFF; padding:0 2px; border:none;" type="submit" value="Query"> 
					   <INPUT name="Submit" style="background:url(images/ioc_anniu_b.gif) no-repeat; color:#FFF; padding:0px; border:none; cursor:pointer;" onclick="window.location='search.asp'" type="button" value="advanced">
						</A></TD>
					</TR>
				</FORM>
			  </TABLE></TD>
		    </TR>
		  </TABLE>
	  </td>
    </tr>
  </tbody>
</table>
