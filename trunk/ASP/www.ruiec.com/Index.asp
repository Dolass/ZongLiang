<!--#include file="Inc/Conn.asp"-->
<%
	if not(blogmode) then
		IF Sdcms_Mode=2 Then
			IF Not Check_File("Index"&Sdcms_Filetxt) Then Index_Html
			Server.Transfer "Index"&Sdcms_Filetxt
		Else
			Index_Html
		End IF
	else
		Index_Blog
	end if	
	Sub Index_Html
		Dim C
		Set C=New sdcms_create
			C.Create_index()
		Set C=Nothing
	End Sub
	
	Sub Index_Blog
		Dim Temp,Show
		Set Temp=New Templates
		DbOpen
		Show=Temp.Sdcms_Load(Load_temp_dir&sdcms_skins_blog)
		Temp.TemplateContent=Show
		Show=Temp.Display
		Temp.Page_Mark(Show)
		
		Dim PageField,PageTable,PageWhere,PageOrder,PagePageSize,PageEof,PageLoop,PageHtml
		PageHtml=Temp.Page_Html
		PageField=Temp.Page_Field
		PageTable=Temp.Page_Table
		PageWhere=Temp.Page_Where
		PageOrder=Temp.Page_Order
		PagePageSize=Temp.Page_PageSize
		PageEof=Temp.Page_Eof
		PageLoop=Temp.Page_Loop
		IF Len(PagePageSize)=0 Then
			Echo "����ȷʹ��ר��ģ��":Died
		End IF
		
		Dim Page:Page=IsNum(Trim(Request.QueryString("Page")),1)
		Dim P,I
		Set P=New Sdcms_Page
		With P
			.Conn=Conn
			.Pagesize=PagePageSize
			.PageNum=Page
			.Table=PageTable
			.Field=PageField
			.Where=PageWhere
			.Key="ID"
			.Order=PageOrder
			Select Case Sdcms_Mode
				Case "1"
					.PageStart="index_"
					.PageEnd=Sdcms_Filetxt
				Case Else
					.PageStart="?Page="
			End Select
		End With
		Set Rs=P.Show
	
		IF Err Then
			Show=Replace(Show,PageHtml,PageEof)
			Temp.Label "{sdcms:listpage}",""
			Err.Clear
		Else
			Dim Get_Loop,t1
			Get_Loop=""
			t1=""
			For I=1 To P.PageSize
				IF Rs.Eof Or Rs.Bof Then Exit For
					t1=PageLoop
					t1=t1&Temp.Get_Page(t1,PageTable,I)
					Get_Loop=Get_Loop&t1
				Rs.MoveNext
			Next
			Get_Loop=Replace(Get_Loop,PageLoop,"")
			Show=Replace(Show,PageHtml,Get_Loop)
			Temp.Label "{sdcms:listpage}",P.PageList
		End IF
		Temp.TemplateContent=Show
		Temp.Analysis_Static()
		Temp.Analysis_Loop()
		Temp.Analysis_IIF()
		Show=Temp.Gzip
		Show=Temp.Display
		Echo Show
		Set Temp=Nothing
	End Sub
%>