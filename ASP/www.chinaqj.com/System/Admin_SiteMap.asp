﻿<% Server.ScriptTimeout=50000 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<%
if Instr(session("AdminPurview"),"|40,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
%>
<%
session("server")="http://"&Request.servervariables("Http_Host")&"" '你的域名
vDir = "/" '制作SiteMap的目录,相对目录(相对于根目录而言)


set objfso = CreateObject("Scripting.FileSystemObject")
root = Server.MapPath(vDir)

'response.ContentType = "text/xml"
'response.write "<?xml version='1.0' encoding='UTF-8'?>"
'response.write "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>"

Function GetGMT(dat)
	dim y,m,d,h,mm,s,r
	week=split("Sun,Mon,Tue,Wed,Thu,Fri,Sat",",")
	mon=split("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",")
	y=year(dat)
	m=mon(month(dat)-1)
	d=day(dat):if d<10 then d="0"&d
	h=hour(dat):if h<10 then h="0"&h
	mm=minute(dat):if mm<10 then mm="0"&mm
	s=second(dat):if s<10 then s="0"&s
	w=week(weekday(dat)-1)
	GetGMT=w&", "&d&" "&m&" "&y&" "&h&":"&mm&":"&s&" GMT"
End Function

chinaqjtime=GetGMT(now)

str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<!-- "
str = str & "Google Site Map File Generated by http://www.chinaqj.com "&chinaqjtime&""
str = str & "-->" & vbcrlf
str = str & "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>" & vbcrlf

Set objFolder = objFSO.GetFolder(root)
'response.write getfilelink(objFolder.Path,objFolder.dateLastModified)
Set colFiles = objFolder.Files
ChinaQJRSS=0
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
.SaveToFile server.mappath("/sitemap.xml"),2 '生成的XML文件名
.Close
End With

Set objStream = Nothing
If Not Err Then
Response.Write("<br />&nbsp;&nbsp;秦江陶瓷为您提供适用于Google的SiteMap。<br />&nbsp;&nbsp;SiteMap.xml生成完毕，共<font color='#FF0000'> "&ChinaQJRSS&" </font>条记录被索引。<br />&nbsp;&nbsp;点击查看 <a href='../sitemap.xml' target='_blank'><font color='#FF0000'><u>SiteMap.xml</u></font></a> <a href='http://www.google.com/webmasters/sitemaps/' target='_blank'><u>立即提交至Google</u></a>")
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
file=replace(file,root,"")
file=replace(file,"\","/")
If FileExtensionIsBad(file) then Exit Function
if month(datafile)<10 then filedatem="0"
if day(datafile)<10 then filedated="0"
filedate=year(datafile)&"-"&filedatem&month(datafile)&"-"&filedated&day(datafile)
getfilelink = "<url><loc>"&server.htmlencode(session("server")&file)&"</loc><lastmod>"&filedate&"</lastmod><changefreq>daily</changefreq><priority>1.0</priority></url>"
Response.Flush
ChinaQJRSS=ChinaQJRSS+1
End Function


Function Folderpermission(pathName)

'需要过滤的目录(不列在SiteMap里面)
PathExclusion=Array("\temp","\_vti_cnf","_vti_pvt","_vti_log","cgi-bin","\admin","\system","\Album","\DatePicker")
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
Extensions = Array("html","htm")
'设置列表的文件名,扩展名不在其中的话SiteMap则不会收录该扩展名的文件

if len(trim(sFileName)) = 0 then
FileExtensionIsBad = true
Exit Function
end if

sFileExtension = right(sFileName, len(sFileName) - instrrev(sFileName, "."))
bFileExtensionIsValid = false 'assume extension is bad
for each sFileExt in extensions
if ucase(sFileExt) = ucase(sFileExtension) then
bFileExtensionIsValid = True
exit for
end if
next
FileExtensionIsBad = not bFileExtensionIsValid
End Function
%>