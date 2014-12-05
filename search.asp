<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title><%=webname%>--Advanced search</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  style="background:url(images/ds_07.jpg); text-align:center;">
<!--#include file="include/head.asp"-->
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="12"></td>
  </tr>
</table>
<table width="961" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="190" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td><table width="190" height="190" border="0" cellpadding="0" cellspacing="0">
          <tr>
              <td align="center" valign="middle" background="images/index_2.gif"><!--#include file="logins.asp"--></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_4.gif" width="190" height="12" alt="" /></td>
      </tr>
      <tr>
        <td background="images/leftbg.jpg"><!--#include file="include/sort.asp"--></td>
      </tr>
      <tr>
        <td><img src="images/leftendbg.jpg" width="190" height="36"></td>
      </tr>
    </table></td>
    <td width="771"  valign="top" bgcolor="#FFFFFF" ><table width="100%" align="center" border="0" cellspacing="0" cellpadding="0"  bordercolor="#CCCCCC">
      <tr>
        <td height="28"  bordercolor="#FFFFFF" background="images/body/pdbg01.gif"><div align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=index.asp><%=webname%></a> >> Commodities advanced search</div></td>
      </tr>
      <tr>
        <td height="260" align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><br>
          <br>
          <br>
          <table width="600" height="224" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td background="images/mingle/serche.gif"><table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr>
                  <td width="159">&nbsp;</td>
                  <td width="381" ><table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
                                          <tr>
                        <td width="29%" align="right"></td>
                        <td width="71%" height="20" ></td>
                      </tr>
					<form name="form2" method="post" action="research.asp">
                      <tr>
                        <td width="29%" align="right">Keyword：</td>
                        <td width="71%" height="30"  style="padding-left:10px"><input  class=wenbenkuang name="searchkey" type="text" id="searchkey">                        </td>
                      </tr>
                      <tr>
                        <td align="right">classify：</td>
                        <td height="30" style="padding-left:10px"><%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * FROM bsort order by anclassidorder",conn,1,1
		  %>
                            <select class="wenbenkuang"  name="anclassid">
                              <option value="0" selected>All classification</option>
                              <%do while not rs.eof%>
                              <option value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%></option>
                              <%rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                            </select></td>
                      </tr>
                      <tr>
                        <td align="right">range of price ：</td>
                        <td height="30" style="padding-left:10px"><select class="wenbenkuang"  name="jiage" id="jiage">
                            <option value="20">20$ below</option>
                            <option value="30">30$ below</option>
                            <option value="50" >50$ below</option>
                            <option value="100">100$ below</option>
                            <option value="500">500$ below</option>
                            <option value="1000">1000$ below</option>
                            <option value="10000">10000$ below</option>
                            <option value="100000" selected>100000$ below</option>
                            <option value="1000000">1000000$ below</option>
                          </select></td>
                      </tr>
                      <tr>
                        <td align="right">Finding Methods：</td>
                        <td height="30" style="padding-left:10px"><script>
	var selects111=[];
<%dim i
	set rs=server.createobject("adodb.recordset")
	rs.open "select * from brand order by pingpaiorder",conn,1,1
	'有大类
%>
	selects111['2']=new Array(
<%
	if rs.recordcount>0 then
	    for i=1 to rs.recordcount
		if i=rs.recordcount then 
%>
		new Option('<%=trim(rs("pingpainame"))%>','<%=trim(rs("pingpainame"))%>'));
<%		else
%>
		new Option('<%=trim(rs("pingpainame"))%>','<%=trim(rs("pingpainame"))%>'),
<%
		end if
		rs.movenext
		next
	else 
%>
		

<%
	end if
rs.close
set rs=nothing
%> 
function showselect()
{
if (document.form2.action.value=="2") {
	T_select.style.display = "";
	//加入内容
		with (document.form2){
			if(action.value) {
				selectname.options.length=0;
				for(var i=0;i<selects111[action.value].length;i++){
					selectname.add(selects111[action.value][i]);
				}
			}
		}
}else{
	T_select.style.display = "none";
} 
}
                  </script>
                            <table border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><select class="wenbenkuang"  name="action" id="action" onChange="showselect();">
                                    <option value="1" selected>Product name</option>
                                    <option value="2">Product brand</option>
                                    <option value="3">Content abstract</option>
                                  </select>                                </td>
                                <td><table border="0" cellspacing="0" cellpadding="0" id=T_select style="DISPLAY: none">
                                    <tr>
                                      <td><select class="wenbenkuang" name="selectname">
                                        </select>                                      </td>
                                    </tr>
                                </table></td>
                              </tr>
                          </table></td>
                      </tr>
                      <tr>
                        <td height="40" colspan="2" align="center"><input class="go-wenbenkuang"  type=submit name="submit" value=" Query ">
                            <input class="go-wenbenkuang"  type=reset name="Clear" value=" Rewrite ">                        </td>
                      </tr>
                    </form>
                  </table></td>
                </tr>
              </table></td>
            </tr>
          </table>          </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2" height="40"></td>
  </tr>
</table>
<!--#include file="include/foot.asp"-->
</body>
</html>
