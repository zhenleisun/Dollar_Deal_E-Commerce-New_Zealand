<marquee onmouseover=this.stop() onmouseout=this.start() direction="up" width="100%" height="85"  scrollamount="2" > 
<table width="80%" border="0" align="left" cellpadding="0" cellspacing="0">

  <tr>
    <td STYLE='PADDING-LEFT:0px'> 
  <%

mess=split(fahuomess,"<br>")
    for i=0 to ubound(mess) %>
	
	<%=trim(mess(i))%>	<hr>
<%next%>
</table>
</marquee>

