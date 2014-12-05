<!-- 您使用的是网趣网上购物系统免费版本，免费版基本功能均可使用，诸多高级功能是不具备的，如果您要网上开店，推荐使用官方正式版，官方网站:http://www.cnhww.com 客服QQ：81447932 / 81447933  -->
<%
dim db
const DatabaseType="ACCESS"     
db="cnhwwdata\cnhww.asp"           
       On Error Resume Next
	dim ConnStr
	dim conn
		ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
		Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open connstr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
		Response.End
	End If
sub CloseConn()
	On Error Resume Next
	If IsObject(Conn) Then
		conn.close
		set conn=nothing
	end if
end sub
%>
