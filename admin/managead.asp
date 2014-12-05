<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if%>
<%if request.QueryString("action")="save" then
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from advertisement",conn,1,3
rs("ad")=trim(request("aad"))
rs("adurl")=trim(request("adurl"))
rs("ad1")=trim(request("aad1"))
rs("ad1url")=trim(request("ad1url"))
rs("ad2")=trim(request("aad2"))
rs("ad2url")=trim(request("ad2url"))
rs("ad3")=trim(request("aad3"))
rs("ad3url")=trim(request("ad3url"))
rs("ad4")=trim(request("aad4"))
rs("ad4url")=trim(request("ad4url"))
rs("ad5")=trim(request("aad5"))
rs("ad5url")=trim(request("ad5url"))
rs("ad6")=trim(request("aad6"))
rs("ad6url")=trim(request("ad6url"))
rs("ad7")=trim(request("aad7"))
rs("ad7url")=trim(request("ad7url"))
rs("ad8")=trim(request("aad8"))
rs("ad8url")=trim(request("ad8url"))

	 rs("pic1")=request.form("pic1")
 	 rs("pic1_lnk")=request.form("pic1_lnk")
     rs("tit1")=request.form("tit1")
 	 
 	 rs("pic2")=request.form("pic2")
 	 rs("pic2_lnk")=request.form("pic2_lnk")
     rs("tit2")=request.form("tit2") 

 	 rs("pic3")=request.form("pic3")
 	 rs("pic3_lnk")=request.form("pic3_lnk")
     rs("tit3")=request.form("tit3")

 	 rs("pic4")=request.form("abcdea")
 	 rs("pic4_lnk")=request.form("pic4_lnk")
     rs("tit4")=request.form("tit4")
	 
  	 rs("pic5")=request.form("abcdef")
 	 rs("pic5_lnk")=request.form("pic5_lnk")
     rs("tit5")=request.form("tit5")
	 
 	 rs("dh")=request.form("dh")
 	 rs("dhurl")=request.form("dhurl")
rs.update
rs.close
set rs=nothing
response.write "<script language=javascript>alert('修改成功！');history.go(-1);</script>"
response.End
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
</head>
<body>

<%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from advertisement",conn,1,1%>
<table class="tableBorder" width="90%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td align="center" background="../images/admin_bg_1.gif" height="25"><b><font color="#ffffff">商城广告设置</font></b></td>
</tr>
<tr> 
<td align="center" valign="top"> 
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF">
        <form name="form1" method="post" action="managead.asp?action=save">
          <tr > 
            <td width="24%" align="center" bgcolor="fbc2c2">广告说明</td>
            <td width="33%" align="center" bgcolor="fbc2c2">图片地址</td>
            <td width="15%" align="center" bgcolor="fbc2c2">链接地址</td>
            <td width="28%" align="center" bgcolor="fbc2c2">链接地址</td>
          </tr>
          <tr align="center" class="tableBorder"> 
            <td  class="forumRowHighlight">滚动图片一( JPG格式)</td>
            <td  class="forumRowHighlight"> <input name="pic1" type="text" id="4582" value="<%=trim(rs("pic1"))%>" size="28" ></td>
            <td  class="forumRowHighlight"> <input name="pic1_lnk" type="text" id="a4" value="<%=trim(rs("pic1_lnk"))%>" size="28" ></td>
            <td  class="forumRowHighlight"><input type="button" name="Submit223" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=pic1&uppath=upfile&&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr align="center" class="tableBorder"> 
            <td  class="forumRowHighlight">滚动图片二( JPG格式)</td>
            <td  class="forumRowHighlight"> <input name="pic2" type="text" id="a" value="<%=trim(rs("pic2"))%>" size="28" ></td>
            <td  class="forumRowHighlight"> <input name="pic2_lnk" type="text" id="a5" value="<%=trim(rs("pic2_lnk"))%>" size="28" ></td>
            <td  class="forumRowHighlight"><input type="button" name="Submit224" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=pic2&uppath=upfile&&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr align="center" class="tableBorder"> 
            <td  class="forumRowHighlight">滚动图片三( JPG格式)</td>
            <td  class="forumRowHighlight"> <input name="pic3" type="text" id="a2" value="<%=trim(rs("pic3"))%>" size="28" ></td>
            <td  class="forumRowHighlight"> <input name="pic3_lnk" type="text" id="a6" value="<%=trim(rs("pic3_lnk"))%>" size="28" ></td>
            <td  class="forumRowHighlight"><input type="button" name="Submit225" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=pic3&uppath=upfile&&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr align="center" class="tableBorder"> 
            <td  class="forumRowHighlight">滚动图片四( JPG格式)</td>
            <td  class="forumRowHighlight"> <input name="abcdea" type="text" id="abcde2" value="<%=trim(rs("pic4"))%>" size="28" ></td>
            <td  class="forumRowHighlight"> <input name="pic4_lnk" type="text" id="a7" value="<%=trim(rs("pic4_lnk"))%>" size="28" ></td>
            <td  class="forumRowHighlight"><input type="button" name="Submit226" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=abcdea&uppath=upfile&&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr align="center" class="tableBorder"> 
            <td  class="forumRowHighlight">滚动图片五( JPG格式)</td>
            <td  class="forumRowHighlight"> <input name="abcdef" type="text" id="abcde" value="<%=trim(rs("pic5"))%>" size="28" ></td>
            <td  class="forumRowHighlight"> <input name="pic5_lnk" type="text" id="pic5_lnk" value="<%=trim(rs("pic5_lnk"))%>" size="28" ></td>
            <td  class="forumRowHighlight"><input type="button" name="Submit2262" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=abcdef&uppath=upfile&&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页中间横幅 734*135 
              </div></td>
            <td align="center" bgcolor="fbf4f4"><input name="aad3" type="text" id="aad3" size="28" value=<%=trim(rs("ad3"))%>></td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad3url" type="text" id="ad3url" size="28" value=<%=trim(rs("ad3url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit227" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad3&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页左侧广告1 190*235</div></td>
            <td align="center" bgcolor="fbf4f4"><input name="aad4" type="text" id="aad4" size="28" value=<%=trim(rs("ad4"))%>></td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad4url" type="text" id="ad4url" size="28" value=<%=trim(rs("ad4url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit228" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad4&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页左侧广告2 190*235 
              </div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad5" type="text" id="aad5" size="28" value=<%=trim(rs("ad5"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad5url" type="text" id="ad5url" size="28" value=<%=trim(rs("ad5url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit229" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad5&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页左侧广告3 190*235 
              </div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad1" type="text" id="aad1" size="28" value=<%=trim(rs("ad1"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad1url" type="text" id="ad1url" size="28" value=<%=trim(rs("ad1url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit225" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad1&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页右侧广告1 180*190</div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad6" type="text" id="aad6" size="28" value=<%=trim(rs("ad6"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad6url" type="text" id="ad6url" size="28" value=<%=trim(rs("ad6url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit2210" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad6&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页右侧广告2 180*190 
              </div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad7" type="text" id="aad7" size="28" value=<%=trim(rs("ad7"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad7url" type="text" id="ad7url" size="28" value=<%=trim(rs("ad7url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit2211" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad7&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">首页右侧广告3 180*190 
              </div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad2" type="text" id="aad2" size="28" value=<%=trim(rs("ad2"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"> <input name="ad2url" type="text" id="ad2url" size="28" value=<%=trim(rs("ad2url"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit226" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad2&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4"><div align="center">分页广告 700*160 
              </div></td>
            <td align="center" bgcolor="fbf4f4"> <input name="aad8" type="text" id="aad8" size="28" value=<%=trim(rs("ad8"))%>> 
            </td>
            <td align="center" bgcolor="fbf4f4"><input name="ad8url" type="text" id="ad8url" size="28" value=<%=trim(rs("ad8url"))%>></td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit2212" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=aad8&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr > 
            <td align="center" bgcolor="fbf4f4">导航条长幅广告</td>
            <td align="center" bgcolor="fbf4f4"><input name="dh" type="text" id="dh" size="28" value=<%=trim(rs("dh"))%>></td>
            <td align="center" bgcolor="fbf4f4"><input name="dhurl" type="text" id="dhurl" size="28" value=<%=trim(rs("dhurl"))%>></td>
            <td align="center" bgcolor="fbf4f4"><input type="button" name="Submit2263" value="广告图片上传" onClick="window.open('../upload.asp?formname=form1&editname=dh&uppath=upfile&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"></td>
          </tr>
          <tr> 
            <td colspan="4"  align="center" bgcolor="fbf4f4">以上广告，的图片尺寸只要宽度尺寸一致即可，高度按页面实际要求！ 
            </td>
          </tr>
          <tr> 
            <td colspan="4"  align="center" bgcolor="fbf4f4"><input type="submit" name="Submit" value="提交更改"></td>
          </tr>
        </form>
      </table>
</td>
</tr>
</table>
</body>
</html>
