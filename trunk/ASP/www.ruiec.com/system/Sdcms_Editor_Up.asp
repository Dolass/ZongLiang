<!--#include file="Sdcms_Check.asp"-->
<!--#include file="../Inc/Upload.asp"-->
<!--#include file="../Inc/AspJpeg.asp"-->
<%
Dim Act:Act=Trim(Request.QueryString("Act"))
Dim Alow_FileType
Select Case Act
	Case "1":Alow_FileType="doc|docx|rar|zip|7z|xls|pdf|ppt|flv"
	Case Else:Act=0:Alow_FileType="jpg|gif|jpeg"
End Select
Const IsOldName=0'0Ϊ�Զ���������1Ϊԭ�ļ���

Dim UpLoad,FileName
Set UpLoad=New UpLoadClass
	UpLoad.FileType=Alow_FileType'�����ϴ����ļ�����
	UpLoad.SavePath=UploadDir("")'�ļ��ϴ���Ŀ¼��������art/����/Ϊ��art�ļ����ڣ����򽫺�����һ����Ϊ�ļ�����
	UpLoad.Charset="gb2312"
	UpLoad.MaxSize=Sdcms_upfileMaxSize'200*1024Ϊ200K
	UpLoad.AutoSave=IsOldName
	UpLoad.Open()
	'�����ļ�
	IF IsOldName=0 Then
		FileName=UpLoad.SavePath&UpLoad.Form("imgFile")
	Else
		FileName=UpLoad.SavePath&UpLoad.Form("imgFile_Name")
	End IF
	Select Case UpLoad.Form("imgFile_Err")
		Case -1:Set_Json 1,"�ļ�Ϊ�գ��ϴ�ʧ��"
		Case  1:Set_Json 1,"�ļ������ϴ�ʧ��"
		Case  2:Set_Json 1,"�������ϴ����ļ�"
		Case  3:Set_Json 1,"�ļ������Ҳ������ϴ���δ������"
		Case  4:Set_Json 1,"�ļ�����ʧ��"
	End Select
	
	Dim File_Type
	Select Case Lcase(Right(Filename,3))
		Case "jpg","gif","png","peg","bmp":File_Type=0
		Case "swf",".rm","wma","mp3","wmv","mvb","avi":File_Type=1
		Case "flv":File_Type=2
		Case Else:File_Type=3
	End Select
	IF File_Type=0 Then
		Sdcms_Jpeg(Filename)
	End IF
	
	Set_Json 0,FileName
Set UpLoad=Nothing

Sub Set_Json(ByVal t0,ByVal t1)
	IF t0=1 Then
		Response.Write("{""error"":"&t0&",""message"":"""&t1&"""}")
		Response.End()
	Else
		Response.Write("{""error"":"&t0&",""url"":"""&t1&"""}")
	End IF
End Sub
%>