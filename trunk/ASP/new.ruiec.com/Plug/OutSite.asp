<!--#include file="../inc/conn.asp"-->
<%
'============================================================
'������ƣ�վ�����
'Website��http://www.sdcms.cn
'Author��ITƽ��
'Date��2010-10-27
'============================================================
Dim ID:ID=IsNum(Request.QueryString("ID"),0)
Get_Info(ID)
CloseDb
Sub Get_Info(ByVal t0)
	DbOpen
	Dim Sql,Rs,Loop_Content,Temp,Show
	Set Rs=Conn.Execute("Select Loop_Content,IsPass,CacheTime From Sd_OutSite Where ID="&t0&"")
	IF Rs.Eof Then
		Echo Loop_Content="Sdcms��ʾ:��������"
	Else
		IF Rs(1)=0 Then
			Loop_Content="Sdcms��ʾ:δͨ�����"
		Else
			Loop_Content=Rs(0)
			Sdcms_CacheDate=Rs(2)'��������ʱ��
		End IF
	End IF
	Rs.Close
	Set Rs=Nothing
	'ǿ���滻��ַ��ǩ
	IF Sdcms_Root<>"/" Then Loop_Content=Replace(Loop_Content,"{sdcms:weburl}",Replace(Sdcms_WebUrl,Sdcms_Root,""))
	
	Sdcms_Cache=True'ǿ�ƻ���
	IF Check_Cache("Sdcms_OutSite_"&t0) Then
		Set Temp=New Templates
		Temp.TemplateContent=Fixjs(Loop_Content)
		Temp.Analysis_Static()
		Temp.Analysis_Loop()
		Temp.Analysis_IIF()
		Show=Info_Re(Temp.Display)
		Set Temp=Nothing
		Create_Cache "Sdcms_OutSite_"&t0,Show
	End IF
	Echo Load_Cache("Sdcms_OutSite_"&t0)
End Sub

Function Fixjs(ByVal t0)
	Dim t1
	IF Len(t0)=0 Then Fixjs="":Exit Function
	t1=Replace(t0,Chr(39),"")
	t1=Replace(t1,Chr(13),"")
	t1=Replace(t1,Chr(10),"")
	t1=Replace(t1,"��","")
	Fixjs=t1
End Function

Function Info_Re(ByVal t0)
	Dim t1
	IF Len(t0)=0 Then Info_Re="":Exit Function
	t1=Replace(t0,Chr(39), "\'")
	't1=Replace(t1,Chr(13), "")
	t1=Replace(t1,"""","\""")
	t1=Replace(t1,"/","\/")
	't1=Replace(t1,vbcrlf,""");"&vbcrlf&"document.writeln(""")
	Info_Re="document.writeln("""&t1&""");"
End Function
%>