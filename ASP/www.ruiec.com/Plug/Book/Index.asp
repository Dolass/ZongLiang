<!--#include file="../../Inc/Conn.asp"-->
<%
'============================================================
'������ƣ�Ajax���Ա�
'Website��http://www.sdcms.cn
'Author��ITƽ��
'Date��2008-11-6
'Update:2010-10
'============================================================
Dim book_pass
Dim Action:Action=Lcase(Trim(Request.QueryString("Action")))
book_pass=0 '�Ƿ���Ҫ��0Ϊ��ˣ�1Ϊֱ��ͨ��
DbOpen
Select Case Action
	Case "save":SaveDb
	Case Else:Show_msg
End select

Sub SaveDb
	IF Check_post Then
		Echo "1��ֹ���ⲿ�ύ����!":Exit Sub
	End IF
	Dim lastpostdate,username,content,content1,yzm
	lastpostdate=Load_Cookies("book")
	lastpostdate=Re(lastpostdate,"&#32;","��")
	IF lastpostdate<>"" then
		IF Int(DateDiff("s",lastpostdate,now()))<=60 then
			Echo "1��������ٶ�̫��!":Died
		End IF
	End IF
	username=Trim(Request.Form("username"))
	content=Request.Form("content")
	content1=Removehtml(Request.Form("content"))
	yzm=Trim(Request.Form("yzm"))
	IF yzm<>Session("SDCMSCode") Then echo "1��֤�����":Exit Sub
	IF username="" Or content="" Then Echo "1����д�ø���ı�":Exit Sub'����Ϊ��
	IF Len(username)<2 Then Echo "1����̫���˰ɣ�":Exit Sub'����̫��
	IF Len(username)>10 Then Echo "1����Ҳ̫���˰�":Exit Sub'����̫��
	IF Len(content1)=0 Then Echo "1���ݱ���������!":Exit Sub'����̫��
	IF Len(content)<5 Then Echo "1��д��ô�����ݣ�":Exit Sub'����̫��
	IF Len(content)>300 Then Echo "1������ݶ����Ҵ治�°���":Exit Sub'����̫��
	username=FilterText(username,1)
	content=FilterHtml(content)
	Dim Rs,Sql
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select username,content,ispass,ip,adddate From sd_book "
	Rs.Open Sql,Conn,1,3
	rs.Addnew
	rs(0)=Left(username,10)
	rs(1)=content
	rs(2)=book_pass
	rs(3)=GetIp
	Rs(4)=Dateadd("h",Sdcms_TimeZone,Now())
	rs.Update
	rs.Close
	Set Rs=Nothing
	IF book_pass=1 then
		Echo "0�����ɹ�"'�����ɹ�
	Else
		Echo "0�����ɹ�����˺���ʾ"'�����ɹ�
	End IF
	Add_Cookies "book",Now()
End Sub

Sub Show_msg
	Sdcms_Cache=False'ǿ�ƹرջ���
	Dim Temp,show
	Set Temp=New Templates
	Show=Temp.Sdcms_Load(Load_temp_dir&sdcms_skins_book)
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
		Echo "����ȷʹ������ģ��":Died
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
		.PageStart="?Page="
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
	Echo Temp.Display
	Set Temp=Nothing
End Sub
%>