<%
'===============================================
'��������GetHttpPage
'��  �ã���ȡ��ҳԴ��
'��  ����t0��ҳ��ַ,t1����
'===============================================
Function GetHttpPage(t0,t1)
Dim Http
On Error Resume Next
IF IsNull(t0) Or Len(t0)<18 Or t0="$False$" Then
	GetHttpPage="$False$"
	Exit Function
End If
BlockStartTime=Timer()
Set Http=Server.Createobject("MSXML2.XMLHTTP")
Http.open "GET",t0,False
Http.Send()
Dim temp,BlockTimeout 	   
BlockTimeout=IsNum(Get_Coll_Configs(0),64)
While(http.ReadyState<>4)
	'�ж��Ƿ�鳬ʱ
	temp=Timer()-BlockStartTime
	IF (temp>BlockTimeout) Then
		http.abort
		Set Http=Nothing 
		GetHttpPage="$False$"
		Exit function
		Died
	End If
	http.waitForResponse 10000'�ȴ�1000����
Wend
IF Http.Readystate<>4 then
	Set Http=Nothing 
	GetHttpPage="$False$"
	Exit Function
End if
GetHTTPPage=ReplaceTrim(bytesToBSTR(Http.responseBody,t1))
Set Http=Nothing
   
IF Err.number<>0 then
	IF IsNull(t0) Or Len(t0)<18 Or t0="$False$" Then
		GetHttpPage="$False$"
		Exit Function
	End If
	Set Http=Nothing
	Err.Clear
End IF 
End Function

'===============================================
'��������BytesToBstr
'��  �ã�����ȡ��Դ��ת��Ϊ����
'��  ����Body ------Ҫת���ı���
'��  ����Cset ------Ҫת��������
'===============================================
Function BytesToBstr(t0,t1)
   Dim Objstream
   Set Objstream=Server.CreateObject("adodb.stream")
   objstream.Type=1
   objstream.Mode=3
   objstream.Open
   objstream.Write t0
   objstream.Position=0
   objstream.Type=2
   objstream.Charset=t1
   BytesToBstr=objstream.ReadText 
   objstream.Close
   Set objstream=Nothing
End Function

'===============================================
'��������GetBody
'��  �ã���ȡ�̶����ַ���
'��  ����t0   ----ԭ�ַ���
'��  ��: t1 ------ ��ʼ�ַ���
'��  ��: t2 ------ �����ַ���
'��  ����t3 ------�Ƿ����StartStr
'��  ����t4 ------�Ƿ����OverStr
'===============================================
Public Function GetBody(t0,t1,t2,t3,t4)
	Dim SS,Match,TempStr,strPattern,s,o  
	IF IsNull(t1) Then GetBody="$False$":Exit Function
	t1=ReplaceTrim(t1):t2=ReplaceTrim(t2)
	s=Len(t1):o=Len(t2)

	IF s=0 Or o=0  Then GetBody="$False$" : Exit Function
	strPattern="("&CorrectPattern(t1)&")(.+?)("&CorrectPattern(t2)&")"
	On Error Resume Next
	Dim res
	Set res=New RegExp
	res.IgnoreCase=False
	res.Global=False
	res.Pattern=strPattern
	Set SS=res.Execute(t0)
	For Each Match In SS
		TempStr=Match.Value
	Next
	IF TempStr="" Then'���ַ���,����������
	   GetBody="$False$"
	   Exit Function
	End If
   
	If t3=False then
		TempStr=Right(TempStr,Len(TempStr)-S)
	End if
	If t4=False then
		TempStr=Left(TempStr,Len(TempStr)-O)
	End if
	If Err.number<>0 then  '����,����������
		GetBody="$False$"
		Exit Function
	End If
	Set SS=Nothing
	Set res=Nothing
	GetBody=TempStr
	Exit Function
End Function

'===============================================
'��������GetArray
'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
'��  ����t0 ------��ȡ��ַ��ԭ�ַ�
'��  ����t1 ------��ʼ�ַ���
'��  ����t2 ------�����ַ���
'��  ����t3 ------�Ƿ����StartStr
'��  ����t4 ------�Ƿ����OverStr
'===============================================
Function GetArray(t0,t1,t2,t3,t4)
   Dim TempStr,TempStr2,objRegExp,Matches,Match,Templisturl,TempStr_i
   Dim s,o 
   On Error Resume Next
   IF t0="$False$" Or t0="" Or IsNull(t0)=True Or t1="" Or t2="" Or  IsNull(t1)=True Or IsNull(t2)=True Then
	  GetArray="$False$"
	  Exit Function
   End IF
   t1=ReplaceTrim(t1) : t2=ReplaceTrim(t2) : t0=t0
   s=Len(t1) : o=Len(t2)
   TempStr=""
   Set objRegExp=New Regexp 
   objRegExp.IgnoreCase=True 
   objRegExp.Global=True
   objRegExp.Pattern="("&CorrectPattern(t1)&").+?("&CorrectPattern(t2)&")"
   Set Matches=objRegExp.Execute(t0) 
   For Each Match in Matches
		TempStr_i=Match.Value
		If t3=False then
			TempStr_i=Right(TempStr_i,Len(TempStr_i) -S)
		End if
		If t4=False then
			TempStr_i=Left(TempStr_i,Len(TempStr_i) - O)
		End if	
		TempStr=TempStr&"$Array$"&TempStr_i
   Next 
   Set Matches=nothing

   IF TempStr="" Then
	  GetArray="$False$"
	  Exit Function
   End IF
   TempStr=Right(TempStr,Len(TempStr)-7)
   Set objRegExp=Nothing
   Set Matches=Nothing

   IF TempStr="" then
	  GetArray="$False$"
   Else
	  GetArray=TempStr
   End IF
End Function

Function ReRepeat(t0)
	Dim t1,ArrData,i,Arrtag,Arrval,t2
	IF Instr(t0,"$Array$")=0 Then ReRepeat=t0:Exit Function
	t1=Split(t0,"$Array$")
	Set ArrData=Server.CreateObject("Scripting.Dictionary")
	For i=0 to Ubound(t1)
		IF ArrData.Exists(t1(i)) Then ArrData.item(t1(i))=t1(i) Else ArrData.Add t1(i),t1(i)
	Next 
	
	Arrtag=ArrData.keys
	Arrval=ArrData.items
	IF ArrData.count>=1 Then
		For i=0 To ArrData.count-1
			IF i>0 then t2=t2&"$Array$"
			t2=t2&Arrval(i)
		Next
	End IF
	ReRepeat=t2
	Set ArrData=Nothing
End Function
	
'===============================================
'��������ReplaceTrim
'��  �ã����˵��ַ������е�tab�ͻس��ͻ���
'===============================================
Function ReplaceTrim(t0)
	Dim res
	Set res=New RegExp
	res.IgnoreCase=True
	res.Global=True
	res.Pattern="(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
	t0=res.Replace(t0, vbNullString)
	Set res=Nothing
	ReplaceTrim=t0
	Exit Function
End Function
	
Function CorrectPattern(t0)
	t0=Re(t0,"\","\\")
	t0=Re(t0,"~","\~")
	t0=Re(t0,"!","\!")
	t0=Re(t0,"@","\@")
	t0=Re(t0,"#","\#")
	t0=Re(t0,"%","\%")
	t0=Re(t0,"^","\^")
	t0=Re(t0,"&","\&")
	t0=Re(t0,"*","\*")
	t0=Re(t0,"(","\(")
	t0=Re(t0,")","\)")
	t0=Re(t0,"-","\-")
	t0=Re(t0,"+","\+")
	t0=Re(t0,"[","\[")
	t0=Re(t0,"]","\]")
	t0=Re(t0,"<","\<")
	t0=Re(t0,">","\>")
	t0=Re(t0,".","\.")
	t0=Re(t0,"/","\/")
	t0=Re(t0,"?","\?")
	t0=Re(t0,"=","\=")
	t0=Re(t0,"|","\|")
	t0=Re(t0,"$","\$")
	CorrectPattern=t0
End Function
	
'================================================
'��������FormatRemoteUrl
'��  �ã���ʽ���ɵ�ǰ��վ������URL-����Ե�ַת��Ϊ���Ե�ַ
'��  ���� t0 ----Url�ַ���
'��  ���� t1 ----��Ȼ��վURL
'����ֵ����ʽ��ȡ���Url
'================================================
Function FormatRemoteUrl(t0,t1)
	Dim strUrl
	IF Len(t0) < 2 Or Len(t0) > 255 Or Len(t1) < 2 Then
		FormatRemoteUrl = vbNullString
		Exit Function
	End IF
	t1 = Trim(Re(Re(Re(Re(Re(t1, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
	t0 = Trim(Re(Re(Re(Re(Re(t0, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))	
	IF InStr(9,t1,"/")=0 Then
		strUrl=t1
	Else
		strUrl=Left(t1,InStr(9,t1,"/")-1)
	End IF

	IF strUrl=vbNullString Then strUrl = t1
	Select Case Left(LCase(t0), 6)
		Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
			FormatRemoteUrl=t0
			Exit Function
	End Select

	IF Left(t0,1)="/" Then
		FormatRemoteUrl=strUrl&t0
		Exit Function
	End IF
	
	IF Left(t0,3)="../" Then
		Dim ArrayUrl
		Dim ArrayCurrentUrl
		Dim ArrayTemp()
		Dim strTemp
		Dim i,n
		Dim c,l
		n=0
		ArrayCurrentUrl=Split(t1, "/")
		ArrayUrl=Split(t0, "../")
		c=UBound(ArrayCurrentUrl)
		l=UBound(ArrayUrl) + 1
		
		IF c>l+2 Then
			For I=0 To c-l
				ReDim Preserve ArrayTemp(n)
				ArrayTemp(n)=ArrayCurrentUrl(I)
				n=n+1
			Next
			strTemp=Join(ArrayTemp, "/")
		Else
			strTemp=strUrl
		End IF
		t0=Re(t0,"../",vbNullString)
		FormatRemoteUrl=strTemp&"/"&t0
		Exit Function
	End If
	strUrl=Left(t1,InStrRev(t1,"/"))
	FormatRemoteUrl=strUrl&Re(t0,"./",vbNullString)
	Exit Function
End Function

Function Reurl(t0,t1)
	Dim res,Matches,Match,t2,PicUrl,i,PicUrl_Old
	Set res=New Regexp 
	   res.IgnoreCase=True 
	   res.Global=True
	   res.Pattern="<img.+?[^\>]>"
	   Set Matches=res.Execute(t0) 
	   For Each Match in Matches
		  IF Len(t2)>0 Then
			 t2=t2&"$Array$"&Match.Value
		  Else
			 t2=Match.Value
		  End if
	   Next
	   
	   IF Len(t2)=0 Then
		   Reurl=t0:Exit Function
	   Else
		  PicUrl=Split(t2,"$Array$")
		  t2=""
		  For i=0 To Ubound(PicUrl)
			 res.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|png)"
			 Set Matches =res.Execute(PicUrl(i)) 
			 For Each Match in Matches
			   IF Len(t2)>0 Then
				   t2=t2&"$Array$"&Match.Value
				Else
				   t2=Match.Value
				End if
			 Next
		  Next
		  'ȥ��������Ϣ
		  Res.Pattern ="src\s*=\s*"
		  t2=Res.Replace(t2,"")
	   End IF
	   
	   'ȥ���ظ�ͼƬ��ʼ
	   PicUrl=Split(t2,"$Array$")
	   t2=""
	   For I=0 To Ubound(PicUrl)
		  If Instr(Lcase(t2),Lcase(PicUrl(i)))<1 Then
			 t2=t2&"$Array$"&PicUrl(i)
		  End If
	   Next
	   t2=Right(t2,Len(t2)-7)
	   PicUrl=Split(t2,"$Array$")
	   'ȥ���ظ�ͼƬ����
	
	   'ת�����ͼƬ��ַ��ʼ
	   t2=""
	   For i=0 To Ubound(PicUrl)
		  t2=t2&"$Array$"&FormatRemoteUrl(PicUrl(i),t1)
	   Next
	   t2=Right(t2,Len(t2)-7)
	   t2=Replace(t2,Chr(0),"")
	   PicUrl_Old=Split(t2,"$Array$")
	   t2=""
	   
	   For i=0 To Ubound(PicUrl_Old)
		   t0=Re(t0,PicUrl(i),PicUrl_Old(i))
	   Next

	Set Res=Nothing
	t0=Re(Lcase(t0),"../http://","http://")
	Reurl=t0
End Function
	
'**************************************************
'��������CreateKeyWord
'��  �ã��ɸ������ַ������ɹؼ���
'��  ����t0---Ҫ���ɹؼ��ֵ�ԭ�ַ���
'����ֵ�����ɵĹؼ���
'**************************************************
Function CreateKeyWord(t0,Num)
   IF t0="" or IsNull(t0)=True or t0="$False$" Then
      CreateKeyWord="$False$"
      Exit Function
   End If
   IF Num="" Or IsNumeric(Num)=False Then
      Num=4
   End If
   t0=Re(t0,CHR(32),"")
   t0=Re(t0,CHR(9),"")
   t0=Re(t0,"&nbsp;","")
   t0=Re(t0," ","")
   t0=Re(t0,"(","")
   t0=Re(t0,")","")
   t0=Re(t0,"<","")
   t0=Re(t0,">","")
   t0=Re(t0,"""","")
   t0=Re(t0,"?","")
   t0=Re(t0,"*","")
   t0=Re(t0,"|","")
   t0=Re(t0,",","")
   t0=Re(t0,".","")
   t0=Re(t0,"/","")
   t0=Re(t0,"\","")
   t0=Re(t0,"-","")
   t0=Re(t0,"@","")
   t0=Re(t0,"#","")
   t0=Re(t0,"$","")
   t0=Re(t0,"%","")
   t0=Re(t0,"&","")
   t0=Re(t0,"+","")
   t0=Re(t0,":","")
   t0=Re(t0,"��","")   
   t0=Re(t0,"��","")
   t0=Re(t0,"��","")
   t0=Re(t0,"��","")
   t0=Re(t0,"&","")  
   t0=Re(t0,"gt;","")      
   Dim ConstrTemp,i
   For i=1 To Len(t0)
      ConstrTemp=ConstrTemp&","&Mid(t0,i,Num)
   Next
   If Len(ConstrTemp)<254 Then
      ConstrTemp=ConstrTemp&","
   Else
      ConstrTemp=Left(ConstrTemp,254)&","
   End If
   ConstrTemp=left(ConstrTemp,Len(ConstrTemp)-1)
   ConstrTemp=Right(ConstrTemp,Len(ConstrTemp)-1)
   CreateKeyWord=ConstrTemp
End Function

't0��Ҫ���˵��ַ�
't1��Ҫ���˵ı�ǩ,��|��
Function Get_Script(t0,t1)	
	Dim t2,i,k
	IF Len(t0)=0 Then Get_Script=t0:Exit Function
	'��ʼ����
	t2=Split(t1,"|")
	For I=0 To Ubound(t2)
		Select Case t2(I)
			Case "Iframe":k=1
			Case "Object","Script"k=2
			Case Else:k=3
		End Select
		t0=ScriptHtml(t0,t2(I),k)
	Next
	Get_Script=t0
End Function

Function ScriptHtml(t0,t1,t2)
    Select Case t2
    Case 1
	   t0=ReplaceText(t0,"<" & t1 & "([^>])*>","")
    Case 2
       t0=ReplaceText(t0,"<" & t1 & "([^>])*>.*?</" & t1 & "([^>])*>","")
	Case 3
		t0=ReplaceText(t0,"<" & t1 & "([^>])*>","")
		t0=ReplaceText(t0,"</" & t1 & "([^>])*>","")
    End Select
    ScriptHtml=t0
End Function

Sub Get_Coll_Replace(t1)
	Dim I,t2
	t1=Split(t1,",")
	For I=0 To Ubound(t1)
		t2=Split(t1(I),"|")
		Content=Re(Content,t2(0),t2(1))
	Next
End Sub

Sub Coll_Filters(t0)
	Dim I,t1,t2
	For I=0 To Ubound(t0,2)
		IF t0(0,I)-ID=0 Then
		
			IF t0(2,I)=1 Then
				IF t0(1,I)=1 Then
					Title=Re(Title,t0(3,I),t0(6,I))
				Else
					Content=Re(Content,t0(3,I),t0(6,I))
				End IF
			Else
				t2=GetBody(t1,t0(4,I),t0(5,I),True,True)
				Do While t2<>"$False$"
					IF t0(1,I)=1 Then
						Title=Re(Title,t2,t0(6,I))
						Title=GetBody(Title,t0(4,I),t0(5,I),True,True)
					Else
						Content=Re(Content,t2,t0(6,I))
						Content=GetBody(Content,t0(4,I),t0(5,I),True,True)
					End IF	
				Loop
			End IF
		End IF
	Next
End Sub

'�ɼ�����

Sub Get_Coll_Config
	IF Sdcms_Cache Then
		IF Check_Cache("Get_Coll_Config") Then
			Create_Cache "Get_Coll_Config",Coll_Config
		End IF
		Get_Coll_Configs=Load_Cache("Get_Coll_Config")
	Else
		Get_Coll_Configs=Coll_Config
	End IF
End Sub

Sub Get_Coll_List(t0)
	IF Sdcms_Cache Then
		IF Check_Cache("Get_Coll_List_"&t0) Then
			Create_Cache "Get_Coll_List_"&t0,Coll_List(t0)
		End IF
		Get_Coll_Lists=Load_Cache("Get_Coll_List_"&t0)
	Else
		Get_Coll_Lists=Coll_List(t0)
	End IF
End Sub

Sub Get_Info_Config(t0)
	IF Sdcms_Cache Then
		IF Check_Cache("Get_Info_Config_"&t0) Then
			Create_Cache "Get_Info_Config_"&t0,Coll_Info_Config(t0)
		End IF
		Get_Info_Configs=Load_Cache("Get_Info_Config_"&t0)
	Else
		Get_Info_Configs=Coll_Info_Config(t0)
	End IF
End Sub

Sub Get_Info_Config(t0)
	IF Sdcms_Cache Then
		IF Check_Cache("Get_Info_Config_"&t0) Then
			Create_Cache "Get_Info_Config_"&t0,Coll_Info_Config(t0)
		End IF
		Get_Info_Configs=Load_Cache("Get_Info_Config_"&t0)
	Else
		Get_Info_Configs=Coll_Info_Config(t0)
	End IF
End Sub

'=============�޻���

Function Coll_Config
	Dim Rs
	Set Rs=Coll_Conn.Execute("Select Timeout,UpfileDir,MaxFileSize,FileExtName From Sd_Coll_Config Where Id=1")
	IF Rs.Eof Then
		Coll_Config=Split("64@@"&Sdcms_Upfiledir&"@@"&Sdcms_upfileMaxSize&"@@"&Sdcms_Upfiletype&"","@@")
	Else
		Coll_Config=Split(""&Rs(0)&"@@"&Rs(1)&"@@"&Rs(2)&"@@"&Rs(3)&"","@@")
	End IF
End Function

'��һ��:��ȡ��Ϣ�б�
Function Coll_List(t0)
	Dim Sql,Rs,ListUrl,ListUrls,i,List_Code,List_Link,List_List,List_Photo,List_Pic,List_Pic_Link,J
	t0=IsNum(t0,0)
	Sql="Select Flag,ListStr,selEncoding,LsString,LoString,HsString,HoString,"'0-6
	Sql=Sql&"ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,LPsString,LPoString,"'7-13,�б��ҳ����
	Sql=Sql&"x_tp,imhstr,imostr,CollecOrder"'14-16,�б�Сͼ,17Ϊ�ɼ�˳��
	Sql=Sql&" From Sd_Coll_Item Where Id="&t0&""
	Set Rs=Coll_Conn.Execute(Sql)
	IF Rs.Eof Then
		Coll_List="û�ҵ�������Ҫ����Ϣ":Response.Flush():Exit Function
	Else
		
		IF Rs(0)=0 Then
			Coll_List="����Ŀδ���ã����ܲɼ�":Response.Flush():Exit Function
		End IF
		IF Rs(7)=0 Then'��������
			ListUrl=Trim(Rs(1))
		ElseIF Rs(7)=1 Then'��������
			ListUrls=Lcase(Rs(8))
			J=1
			IF Clng(Rs(9))>Clng(Rs(10)) Then
				J=-1
			End IF
			For I=Clng(Rs(9)) To Clng(Rs(10)) step J
				ListUrl=ListUrl&Re(ListUrls,"{$id}",I)
				IF I<>Clng(Rs(10)) Then
					ListUrl=ListUrl&"|"
				End IF
			Next
		ElseIF Rs(7)=2 Then'�ֶ����
			ListUrl=Trim(Rs(11))
		End IF
		ListUrl=Split(ListUrl,"|")
		For I=0 To Ubound(ListUrl)
			List_Code=List_Code&GetHttpPage(ListUrl(I),Rs(2))'��ȡ�б�ҳ����
			Coll_Msg List_Code,"��ȡ��ҳԴ��",ListUrl(I)
			
			List_Code=GetBody(List_Code,Rs(3),Rs(4),False,False)'��ȡ�޳��������ݵ��б�
			Coll_Msg List_Code,"��ȡ��Ϣ�б�",ListUrl(I)
			
			IF Rs(14)=1 Then
				List_Photo=List_Photo&GetArray(List_Code,Rs(15),Rs(16),False,False)'��ȡ������Ϣ��Сͼ
				Coll_Msg List_Photo,"��ȡ��ϢСͼ",ListUrl(I)
				IF I<Ubound(ListUrl) Then
					List_Photo=List_Photo&"$Array$"
				End IF
			End IF

			List_Link=List_Link&GetArray(List_Code,Rs(5),Rs(6),False,False)'��ȡ����·����Url
			Coll_Msg List_Code,"��ȡ��Ϣ����",ListUrl(I)
			IF I<Ubound(ListUrl) Then
				List_Link=List_Link&"$Array$"
			End IF
		
		Next
		'Echo List_Link:died
		List_List=Split(List_Link,"$Array$")
		List_Link=Split(List_Link,"$Array$")
		'�������
		IF Rs(17)=1 Then
			For I=Ubound(List_List) To 0 step -1
				List_List(I)=Trim(FormatRemoteUrl(List_Link(Abs(I-Ubound(List_List))),Rs(1)))'�����·��ת��Ϊ����·��
			Next
		Else
			For I=0 to Ubound(List_List)
				List_List(I)=Trim(FormatRemoteUrl(List_List(I),Rs(1)))'�����·��ת��Ϊ����·��
			Next
		End IF
		
		IF Len(List_Photo)>0 Then
			List_Pic_Link=Split(List_Photo,"$Array$")
			List_Pic=Split(List_Photo,"$Array$")
			
			'�������
			IF Rs(17)=1 Then
				For I=Ubound(List_Pic) To 0 step -1
				List_Pic(I)=Trim(FormatRemoteUrl(List_Pic_Link(Abs(I-Ubound(List_Pic))),Rs(1)))'�����·��ת��Ϊ����·��
				Next
			Else
				For I=0 to Ubound(List_Pic)
					List_Pic(I)=Trim(FormatRemoteUrl(List_Pic(I),Rs(1)))'�����·��ת��Ϊ����·��
				Next
			End If

			IF Sdcms_Cache Then
				IF Check_Cache("Coll_Pic_List_"&t0) Then
					Create_Cache "Coll_Pic_List_"&t0,List_Pic
				End IF
			End IF
			
		End IF
		
		Coll_List=List_List
	End IF
End Function

Function  Coll_Info_Config(t0)
	Dim Sql,Rs,Rs_c
	Sql="Select Flag,selEncoding,TsString,ToString,CsString,CoString,AuthorType,AsString,AoString,AuthorStr,"'0-9
	Sql=Sql&"CopyFromType,FsString,FoString,CopyFromStr,DateType,DsString,DoString,KeyType,KsString,KoString,KeyStr,"'10-20
	Sql=Sql&"NewsPaingType,NPsString,NPoString,NewsUrlPaing_s,NewsUrlPaing_o,CollecNewsNum,"'21-25,���ݷ�ҳ����,26Ϊ��������
	Sql=Sql&"Classid,SpecialID,Passed,SaveFiles,Thumb_WaterMark,Coll_Top,Hits,"'27-33
	Sql=Sql&"Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_table,Script_tr,Script_Span,Script_Img,Script_Font,"'34-43
	Sql=Sql&"Script_A,Script_Html,Script_Td,strReplace,ListStr"'44-48
	Sql=Sql&" From Sd_Coll_Item Where Id="&t0&""
	Set Rs=Coll_Conn.Execute(Sql)
	IF Rs.Eof Then
		Echo "û�ҵ�������Ҫ����Ϣ":Response.Flush():Exit Function
	Else
		IF Rs(0)=0 Then
			Echo "����Ŀδ���ã����ܲɼ�":Response.Flush():Exit Function
		End IF
		Set Rs_c=Conn.Execute("Select count(id) from sd_class where id="&rs(27)&"")
		IF rs_c(0)=0 Then
			Echo "<b>ʧ����Ϣ��</b>����������������Ŀ���ã�<br>":Died
		End IF
		IF rs(28)>0 Then
			Set Rs_c=Conn.Execute("Select count(id) from sd_topic where id="&rs(28)&"")
			IF rs_c(0)=0 Then
				Echo "<b>ʧ����Ϣ��</b>ר���������������Ŀ���ã�":Died
			End IF
		End IF
		Coll_Info_Config=Rs.GetRows
	End IF
End Function

'��ȡĿ����ַ����
Function  Get_Url_Content(t0)
	Dim Rs,Sql
	Collection_Data
	Sql="select ListStr,selEncoding From Sd_Coll_Item Where ID="&t0&""
	Set Rs=Coll_Conn.Execute(Sql)
	IF Rs.Eof Then
		Get_Url_Content="����ʧ��"
	Else
		Get_Url_Content=GetHttpPage(Rs(0),Rs(1))
	End IF
End Function

'��ȡĿ����ַ
Function Get_Urls(t0)
	Dim Sql,Rs,ListUrl,J,I,List_Code,List_List,List_Link,ListUrls
	Collection_Data
	Sql="Select Flag,ListStr,selEncoding,LsString,LoString,HsString,HoString,"'0-6
	Sql=Sql&"ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,LPsString,LPoString,"'7-13,�б��ҳ����
	Sql=Sql&"x_tp,imhstr,imostr,CollecOrder"'14-16,�б�Сͼ,17Ϊ�ɼ�˳��
	Sql=Sql&" From Sd_Coll_Item Where Id="&t0&""
	Set Rs=Coll_Conn.Execute(Sql)
	IF Rs.Eof Then
		Echo "û�ҵ�������Ҫ����Ϣ":Response.Flush():Exit Function
	Else
		IF Rs(7)=0 Then'��������
			ListUrl=Trim(Rs(1))
		ElseIF Rs(7)=1 Then'��������
			ListUrls=Lcase(Rs(8))
			J=1
			IF Clng(Rs(9))>Clng(Rs(10)) Then
				J=-1
			End IF
			For I=Clng(Rs(9)) To Clng(Rs(10)) step J
				ListUrl=ListUrl&Re(ListUrls,"{$id}",I)
				IF I<>Clng(Rs(10)) Then
					ListUrl=ListUrl&"|"
				End IF
			Next
		ElseIF Rs(7)=2 Then'�ֶ����
			ListUrl=Trim(Rs(11))
		End IF
		ListUrl=Split(ListUrl,"|")
		For I=0 To Ubound(ListUrl)
			List_Code=List_Code&GetHttpPage(ListUrl(I),Rs(2))'��ȡ�б�ҳ����
			
			List_Code=GetBody(List_Code,Rs(3),Rs(4),False,False)'��ȡ�޳��������ݵ��б�

			List_Link=List_Link&GetArray(List_Code,Rs(5),Rs(6),False,False)'��ȡ����·����Url
			IF I<Ubound(ListUrl) Then
				List_Link=List_Link&"$Array$"
			End IF
		
		Next
		List_List=Split(List_Link,"$Array$")
		List_Link=Split(List_Link,"$Array$")
		'�������
	 
		For I=0 to Ubound(List_List)
			List_List(I)=Trim(FormatRemoteUrl(List_List(I),Rs(1)))'�����·��ת��Ϊ����·��
			Echo "<a href="""&List_List(I)&""" target=""_blank"">"&List_List(I)&"</a><br>"&vbcrlf
			Response.Flush()
		Next

	End IF

End Function
%>