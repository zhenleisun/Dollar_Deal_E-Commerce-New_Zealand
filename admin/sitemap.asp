<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else

end if%>
<%dim count
set rs=server.createobject("adodb.recordset")
rs.open "select weburl  from  webinfo  ",conn,1,1
Server.ScriptTimeout=50000
session("server")=rs("weburl")               '你的域名
vDir = "../"                                               '制作SiteMap的目录,相对目录(相对于根目录而言)
set objfso = CreateObject("Scripting.FileSystemObject")
root = Server.MapPath(vDir)

'response.ContentType = "text/xml"
'response.write "<?xml version='1.0' encoding='UTF-8'?>"
'response.write "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>"

str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>" & vbcrlf

Set objFolder = objFSO.GetFolder(root)
'response.write getfilelink(objFolder.Path,objFolder.dateLastModified)
Set colFiles = objFolder.Files
For Each objFile In colFiles
        'response.write getfilelink(objFile.Path,objfile.dateLastModified)
        str = str & getfilelink(objFile.Path,objfile.dateLastModified) & vbcrlf
Next
ShowSubFolders(objFolder)

'response.write "</urlset>"
str = str & "</urlset>" & vbcrlf
set fso = nothing

Set objStream = Server.CreateObject("ADODB.Stream")
    With objStream
    '.Type = adTypeText
    '.Mode = adModeReadWrite
    .Open
    .Charset = "utf-8"
    .Position = objStream.Size
    .WriteText=str
    .SaveToFile server.mappath("../sitemap.xml"),2 '生成的XML文件名
    .Close
    End With

  Set objStream = Nothing
  If Not Err Then
    Response.Write("<script>alert('成功生成站点地图!');history.back();</script>")
    Response.End
  End If

Sub ShowSubFolders(objFolder)
        Set colFolders = objFolder.SubFolders
        For Each objSubFolder In colFolders
                if folderpermission(objSubFolder.Path) then
                        'response.write getfilelink(objSubFolder.Path,objSubFolder.dateLastModified)
                        str = str & getfilelink(objSubFolder.Path,objSubFolder.dateLastModified) & vbcrlf
                        Set colFiles = objSubFolder.Files
                        For Each objFile In colFiles
                                'response.write getfilelink(objFile.Path,objFile.dateLastModified)
                                str = str & getfilelink(objFile.Path,objFile.dateLastModified) & vbcrlf
                        Next
                        ShowSubFolders(objSubFolder)
                end if
        Next
End Sub


Function getfilelink(file,datafile)
        file=replace(file,root,"",1,-1,1)
        file=replace(file,"\","/")
        If FileExtensionIsBad(file) then Exit Function
        if month(datafile)<10 then filedatem="0"
        if day(datafile)<10 then filedated="0"
        filedate=year(datafile)&"-"&filedatem&month(datafile)&"-"&filedated&day(datafile)
        getfilelink = "<url><loc>"&server.htmlencode(session("server")&file)&"</loc><lastmod>"&filedate&"</lastmod><changefreq>daily</changefreq><priority>1.0</priority></url>"
        Response.Flush
End Function


Function Folderpermission(pathName)

        '需要过滤的目录(不列在SiteMap里面)
        PathExclusion=Array("\blog","\temp","\_vti_cnf","_vti_pvt","_vti_log","cgi-bin","\admin","\edu")
        Folderpermission =True
        for each PathExcluded in PathExclusion
                if instr(ucase(pathName),ucase(PathExcluded))>0 then
                        Folderpermission = False
                        exit for
                end if
        next
End Function


Function FileExtensionIsBad(sFileName)
        Dim sFileExtension, bFileExtensionIsValid, sFileExt
        'modify for your file extension (http://www.googleguide.com/file_type.html)
        Extensions = Array("asp","png","jpeg","zip","pdf","ps","html","htm","php","wk1","wk2","wk3","wk4","wk5","wki","wks","wku","lwp","mw","xls","ppt","doc","wks","wps","wdb","wri","rtf","ans","txt")
'设置列表的文件名,扩展名不在其中的话SiteMap则不会收录该扩展名的文件

        if len(trim(sFileName)) = 0 then
                FileExtensionIsBad = true
                Exit Function
        end if

        sFileExtension = right(sFileName, len(sFileName) - instrrev(sFileName, "."))
        bFileExtensionIsValid = false        'assume extension is bad
        for each sFileExt in extensions
                if ucase(sFileExt) = ucase(sFileExtension) then
                        bFileExtensionIsValid = True
                        exit for
                end if
        next
        FileExtensionIsBad = not bFileExtensionIsValid
End Function
%>
