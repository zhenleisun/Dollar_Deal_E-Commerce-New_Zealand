<table width="189" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/body/newtop.gif" width="189" height="40" border="0" usemap="#Map" /></td>
  </tr>
  <tr>
    <td background="images/body/newbg.gif"  >
      <table  width="98%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <%dim i
		  i=0
		  set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 6 newsname,adddate,newsid from news order by adddate desc",conn,1,1
		  if rs.eof and rs.bof then
		  response.write "目前还没有商城新闻"
		  else
		  do while not rs.eof
		  i=i+1%>
          <td width="10%"  valign="middle">
          <div align="center"><img src="images/body/blit_01.gif" height="3" width="2" /></div></td>
          <td width="90%"height="18" valign="middle"><span class="noti_text">
            <%if len(trim(rs("newsname")))>11 then
		  response.write "<a href=""trends.asp?id="&rs("newsid")&""" title="&year(rs("adddate"))&"年"&month(rs("adddate"))&"月"&day(rs("adddate"))&"日发布>"&left(trim(rs("newsname")),11)&".."&"</a><br>"
		  else
		  response.write "<a href=""trends.asp?id="&rs("newsid")&""" title="&year(rs("adddate"))&"年"&month(rs("adddate"))&"月"&day(rs("adddate"))&"日发布>"&trim(rs("newsname"))&"</a>"
		  end if%>
          </span></td>
        </tr>
        <%rs.movenext
		  loop
		  end if
		  rs.close
		  set rs=nothing%>
      </table></td>
  </tr>
  <tr>
    <td ><img src="images/body/newbot.gif" width="189" height="14" /></td>
  </tr>
</table>
<map name="Map" id="Map">
  <area shape="rect" coords="39,8,117,29" href="trend.asp" />
</map>