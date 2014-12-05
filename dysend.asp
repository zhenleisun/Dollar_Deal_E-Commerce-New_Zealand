<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="#">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
<title>邮件订阅</title>
<SCRIPT LANGUAGE="JavaScript">
function checkmail(){
if(document.youjian.useremail.value.length!=0)
  {
    if (document.youjian.useremail.value.charAt(0)=="." ||        
         document.youjian.useremail.value.charAt(0)=="@"||       
         document.youjian.useremail.value.indexOf('@', 0) == -1 || 
         document.youjian.useremail.value.indexOf('.', 0) == -1 || 
         document.youjian.useremail.value.lastIndexOf("@")==document.youjian.useremail.value.length-1 || 
         document.youjian.useremail.value.lastIndexOf(".")==document.youjian.useremail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.youjian.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.youjian.useremail.focus();
   return false;
   }
}
</script>
</head>
<body>
<%
email=request("useremail")
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from dyemail ",conn,1,3
rs.addnew
rs("email")=email
rs.update
response.Write "<script language=javascript>alert('订阅成功！');history.go(-1);</script>"
response.End
%>
</body>
</html>