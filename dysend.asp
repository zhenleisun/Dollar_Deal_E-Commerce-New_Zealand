<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="#">
<meta name="keywords" content="">
<link href="images/css.css" rel="stylesheet" type="text/css">
<title>�ʼ�����</title>
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
      alert("Email��ַ��ʽ����ȷ��");
      document.youjian.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email����Ϊ�գ�");
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
response.Write "<script language=javascript>alert('���ĳɹ���');history.go(-1);</script>"
response.End
%>
</body>
</html>