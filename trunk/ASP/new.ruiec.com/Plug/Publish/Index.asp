<!--#include file="../../Inc/Conn.Asp"-->
<%
'============================================================
'������ƣ�Ͷ����
'Website��http://www.sdcms.cn
'Author��ITƽ��
'Date��2009-6-14
'LastUpDate:2011-2
'============================================================
	Dim Publish_Action,Publish_Content_Len
	Dim Action:Action=Lcase(Trim(Request.QueryString("Action")))
	Publish_Action=True '����ΪTrue �ر�ΪFalse
	Publish_Content_Len=5000 'Ͷ��������󳤶�(��λ���ַ�)
	IF Not Publish_Action Then
		Echo "0ϵͳ�ѹر�����Ͷ�幦�ܣ�":Died
	End IF
	Sub SaveInfo
		IF Check_post Then
			 Echo "0��ֹ���ⲿ�ύ����":Exit Sub
		End IF
		Dim LastPostDate,t0,t1,t2,t3,t4,t5,t6,t7,content1,Rs,Sql
		LastPostDate=Load_Cookies("Infoadddate")
		IF LastPostDate<>"" Then
			IF Int(DateDiff("s",LastPostDate,Now()))<=60 then
				Echo "0�ύ̫�죬���������":Exit Sub
			End IF
		End IF
		t0=FilterText(Trim(Request.Form("t0")),1)
		t1=FilterText(Trim(Request.Form("t1")),1)
		t2=FilterText(Trim(Request.Form("t2")),1)
		t3=IsNum(Trim(Request.Form("t3")),0)
		t4=Trim(Request.Form("t4"))
		t5=FilterText(Trim(Request.Form("t5")),1)
		t6=FilterHtml(Request.Form("t6"))
		content1=Removehtml(Request.Form("t6"))
		t7=FilterText(Trim(Request.Form("t7")),1)
		t4=Re(t4,"  "," ")
		t4=Re(t4,"��",",")
		t4=Re(t4," ",",")
		t4=Re(t4,"��",",")
		t4=Check_Event(t4,",","")
		IF t7<>Session("SDCMSCode") Then Echo "0��֤�����":Exit Sub
		IF t0="" Or t1="" Or t2="" Or t3="" Or t6="" Then Echo "0���ݲ�����":Exit Sub
		IF Len(t0)<2 Then Echo "0����̫��":Exit Sub
		IF Len(t0)>50 Then Echo "0����Ҳ̫���˰�":Exit Sub
		IF Len(t1)<2 Then Echo "0����̫��":Exit Sub
		IF Len(t1)>10 Then Echo "0����̫��":Exit Sub
		IF t3=0 Then Echo "0���ѡ�����":Exit Sub
		IF Len(content1)=0 Then Echo "0Ͷ�������б��������֣�":Exit Sub
		IF Len(content1)<50 Then Echo "0����̫�̲�����Ͷ�壡":Exit Sub
		DbOpen
		Set Rs=Conn.Execute("Select Class_Type From Sd_Class Where ID="&t3&"")
		IF Rs.Eof Then
			Echo "0û��˼��,Ϲ���ڣ�":Exit Sub
		End IF
		Rs.Close:Set Rs=Nothing
	 
		Set Rs=Server.CreateObject("Adodb.RecordSet")
		Sql="Select title,author,comefrom,classid,content,lastupdate,iscomment,userid,jj,LikeIDType,LikeID,tags,id,htmlname,adddate From sd_info where title Like '%"&t0&"%'"
		Rs.Open Sql,Conn,1,3
		IF Not Rs.Eof Then
			Echo "0ϵͳ�Ѵ�����Ͷ�ݵ����±���":Exit Sub
		Else
			Rs.Addnew
			Rs(0)=Left(t0,255)
			Rs(1)=Left(t1,50)
			Rs(2)=Left(t2,50)
			Rs(3)=t3
			Rs(4)=Left(t6,Publish_Content_Len)
			Rs(5)=Dateadd("h",Sdcms_TimeZone,Now())
			Rs(6)=Sdcms_Comment_Pass
			Rs(7)=-1
			IF Len(t5)=0 Then
				Rs(8)=CloseHtml(CutStr(Content_Decode(Re_Html(t6)),Sdcms_Length))
				
			Else
				Rs(8)=CloseHtml(Content_Decode(t5))
			End IF
			Rs(9)=0
			Rs(10)=0
			Rs(11)=t4
			Rs(13)=Left(Re_filename(Sdcms_Filename),50)
			Rs(14)=Dateadd("h",Sdcms_TimeZone,Now())
			rs.Update
			Rs.MoveLast
			Dim ID
			ID=Rs(12)
			Custom_HtmlName Rs(13),"sd_info",t0,id
			Add_tags(t4)
			Rs.Close:Set Rs=Nothing
			Echo "1Ͷ��ɹ�����ȴ����!"
			Add_Cookies "Infoadddate",Now()
		End IF
	End Sub
	
	Select Case Action
		Case "save":SaveInfo
		Case Else
			Dim Temp,Show
			Set Temp=New Templates
			Temp.Label "{sdcms:class_id}",0
			Temp.Label "{sdcms:class_title}","����Ͷ��"
			Temp.Label "{sdcms:class_title_site}","����Ͷ��"
			Temp.Label "{sdcms:map_nav}",""
			'Temp.Load(Load_temp_dir&sdcms_skins_Publish)
			'Echo Temp.Display
			Show=Temp.Sdcms_Load(Sdcms_Root&"skins/"&Sdcms_Skins_Root&"/"&sdcms_skins_Publish)
		
			Temp.TemplateContent=Show
			Temp.Analysis_Static()
			Show=Temp.Display
			Temp.Page_Mark(Show)
			
			Temp.TemplateContent=Show
			Temp.Analysis_Static()
			Temp.Analysis_Loop()
			Temp.Analysis_IIF()
			Show=Temp.Gzip
			Show=Temp.Display

			Echo Show

			Set Temp=Nothing
	End Select
%>