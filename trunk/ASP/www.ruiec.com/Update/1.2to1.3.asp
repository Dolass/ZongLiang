<!--#include file="../inc/conn.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.cachecontrol = "no-cache"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>SDCMS���ݿ���������</title>
<link href="images/sdcms.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div class="width">
	<div class="left">
	<div id="update_img"><img src="images/update.gif" alt="SDCMS��վ��Ϣ����ϵͳ"/>
	  <div>
	<span class="title">�ٷ���վ</span>��<a href="Http://www.sdcms.cn" target="_blank">Http://www.sdcms.cn</a><br><span class="title">�ٷ���̳</span>��<a href="Http://bbs.sdcms.cn" target="_blank">Http://bbs.sdcms.cn</a></div>
	</div>
    </div>
	
	<div class="left">
		<h1>SDCMS���ݿ���������</h1>
		<dl>
			<dt class="update_content"><strong>����˵����</strong>
				<ul>
					<li class="action">��������SDCMS1.2Sp1 To SDCMS1.3Beta��</li>
					<li>����ǰ���ȱ��ݺ�ԭ��վ�������ݣ�����������⣻</li>
					<li>��������������رձ�ҳ�棬��������������⣻</li>
					<li class="action">��������ɾ�����������򣬲�������˵������������</li>
				</ul>
			</dt>
			<dt><input class="bnt" id="load_bnt" type="button" value="��ʼ����" onClick="if(confirm('׼������������?'))location.href='?action=update';return false;" /><div id="load_update"><div id="load_updates"></div></div></dt>	
		</dl>
	</div>
	
</div>
<%
Select Case action
	Case "update":update_sdcms
End Select

Sub Progress(t0)
	Echo "<script>document.getElementById(""load_update"").style.display='block';document.getElementById(""load_updates"").style.width="""&t0&"%"";document.getElementById(""load_updates"").innerHTML="""&t0&"%"";</script>"
	Response.Flush()
End Sub

Sub update_sdcms
'=========================================
'˵����SDCMS�����ļ�
'���ã�1.2sp1 To 1.3beta
'Author��ITƽ��
'Date��2010-11-28
'=========================================
	Echo "<script>document.getElementById(""load_bnt"").disabled=true;</script>"
	On Error ReSume Next
	IF Sdcms_DataType Then SqlNowString="Now()" Else SqlNowString="Getdate()"
	Set Conn=Server.CreateObject("ADODB.Connection")
	Conn.open ConnStr

	Sql="CREATE TABLE [Sd_OutSite] ("
	Sql=Sql&"	[ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL ,"
	Sql=Sql&"	[Title] nvarchar (50) NULL ,"
	Sql=Sql&"	[CacheTime] int NULL default 0,"
	Sql=Sql&"	[Loop_Content]  ntext NULL ,"
	Sql=Sql&"	[ispass] int NULL default 0,"
	Sql=Sql&"	[AddDate] smalldatetime NULL default "&SqlNowString
	Sql=Sql&")"
	Conn.Execute(Sql)
	Progress 5

	Sql="Alter Table [Sd_Class] Drop class_num,filename"
	Conn.Execute(Sql)
	Progress 10
	
	IF Sdcms_DataType Then
		Set Cat=Server.CreateObject("ADOX.Catalog")
			Cat.ActiveConnection=Connstr
			Cat.Tables("Sd_Class").Columns("classdir")="ClassUrl"
		Set Cat=Nothing
	Else
		Sql="EXEC sp_rename 'Sd_Class.classdir','ClassUrl','COLUMN'"
		Conn.Execute(Sql)
	End IF
	Progress 15
	
	Sql="Alter Table [Sd_Class] add Depth int Default 0"
	Conn.Execute(Sql)
	Progress 20
	
	'��Ҫ���¼�����������depth��ֵ
	Set Rs_Depth=Conn.Execute("Select ID,Partentid From Sd_Class")
	DbQuery=DbQuery+1
	While Not Rs_Depth.Eof
		Depth=Ubound(Split(Rs_Depth(1),","))+1
		Conn.Execute("Update Sd_Class Set Depth="&Depth&" Where Id="&Rs_Depth(0)&"")
		DbQuery=DbQuery+1
	Rs_Depth.MoveNext
	Wend
	Rs_Depth.Close
	Set Rs_Depth=Nothing
	Progress 25
	
	Sql="CREATE View [View_Info] as "
	Sql=Sql&" SELECT Sd_Info.*,Sd_Class.Title AS ClassName,Sd_Class.ClassUrl AS ClassUrl,Sd_Class.FollowID AS FollowID,"
	Sql=Sql&"Sd_Class.PartentID AS PartentID,Sd_Class.Show_Temp AS Show_Temp"
	Sql=Sql&" FROM Sd_Info LEFT OUTER JOIN Sd_Class ON Sd_Info.ClassID=Sd_Class.ID"
	Conn.Execute(Sql)
	Progress 30
	
	Sql="Alter Table [Sd_Info] Drop iscreate,IsSdcms"
	Conn.Execute(Sql)
	Progress 35
	
	Sql="Alter Table [Sd_Info] add Style nvarchar(255) Null"
	Conn.Execute(Sql)
	Progress 40
	
	Sql="Drop Table [sd_skins]"
	Conn.Execute(Sql)
	Progress 45
	Sql="Alter Table [Sd_Comment] add followid int default 0,SupportNum int default 0,allfollowid ntext Null"
	Conn.Execute(Sql)
	Progress 50
	'��Ҫ���������ֶε�ֵ
	Conn.Execute("Update Sd_Comment Set followid=0,SupportNum=0")
	Progress 55
	Sql="Drop Table [sd_notice]"
	Conn.Execute(Sql)
	Progress 60
	Sql="Drop Table [Sd_Count]"
	Conn.Execute(Sql)
	Progress 70
	Sql="Alter Table [Sd_Other] add Page_Key nvarchar(255) Null,Page_Desc ntext Null"
	Conn.Execute(Sql)
	Progress 75
	Sql="Alter Table [Sd_Topic] Drop Root,Filename"
	Conn.Execute(Sql)
	Progress 80
	
	Progress 85
	Conn.Execute("Update Sd_Class Set channel_temp='',list_temp='',show_temp=''")
	
	Progress 90
	Conn.Execute("Update Sd_Other Set page_temp=''")
	
	Progress 95
	Conn.Execute("Update Sd_Topic Set Temp_Dir=''")
	
	'������ϵͳ����
	Config=""
	Config=Config&"<"
	Config=Config&"%"& vbcrlf
	Config=Config&"'��վ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_WebName:Sdcms_WebName="""&Sdcms_WebName&""""&vbcrlf
	Config=Config&"'��վ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_WebUrl:Sdcms_WebUrl="""&Sdcms_WebUrl&""""&vbcrlf
	Config=Config&"'ϵͳĿ¼����Ŀ¼Ϊ��""/"",����Ŀ¼��ʽΪ��""/sdcms/"""& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Root:Sdcms_Root="""&Sdcms_Root&""""&vbcrlf
	Config=Config&"'����ģʽ,0Ϊ��̬Ĭ�ϣ�1Ϊα��̬��2Ϊ��̬"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Mode:Sdcms_Mode=2"&vbcrlf
	Config=Config&"'����Ŀ¼����Ŀ¼Ϊ��,Ҳ����ָ��ΪĳһĿ¼����ʽΪ��""sdcms/"",�����ԣ�""/""����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_HtmDir:Sdcms_HtmDir="""&Sdcms_HtmDir&""""&vbcrlf
	Config=Config&"'ʱ������,�벻Ҫ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_TimeZone:Sdcms_TimeZone=0"&vbcrlf
	Config=Config&"'�Ƿ���������־"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_AdminLog:Sdcms_AdminLog="&Sdcms_AdminLog&""&vbcrlf
	Config=Config&"'�Ƿ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cache:Sdcms_Cache="&Sdcms_Cache&""&vbcrlf
	Config=Config&"'����ǰ׺,����ж��վ�㣬�����ò�ͬ��ֵ"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cookies:Sdcms_Cookies="""&Sdcms_Cookies&""""&vbcrlf
	Config=Config&"'����ʱ�䣬��λΪ��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_CacheDate:Sdcms_CacheDate=60"&vbcrlf
	Config=Config&"'ϵͳ�ļ��ĺ�׺�������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileTxt:Sdcms_FileTxt="".Html"""&vbcrlf
	Config=Config&"'ϵͳģ��Ŀ¼,һ�㲻�����ֶ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Skins_Root:Sdcms_Skins_Root=""2009"""&vbcrlf
	Config=Config&"'ϵͳ�����ļ���Ĭ���ļ��������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileName:Sdcms_FileName=""{ID}"""&vbcrlf
	Config=Config&"'����Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileDir:Sdcms_UpfileDir="""&Sdcms_UpfileDir&""""&vbcrlf
	Config=Config&"'�����ϴ��ļ�����,�������|��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileType:Sdcms_UpfileType="""&Sdcms_UpfileType&""""&vbcrlf
	Config=Config&"'�����ϴ����ļ����ֵ����λ��KB"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_upfileMaxSize:Sdcms_upfileMaxSize="&Sdcms_upfileMaxSize&""&vbcrlf
	Config=Config&"'����ƫ������,ϵͳ�Զ����ɣ���������Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Create_Set:Sdcms_Create_Set=""0, 1, 2"""&vbcrlf
	Config=Config&"'����GOOGLE��ͼ�Ĳ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Create_GoogleMap:Sdcms_Create_GoogleMap=Split(""20,daily,0.8"","","")"&vbcrlf
	Config=Config&"'���ۿ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Comment_Pass:Sdcms_Comment_Pass="&Sdcms_Comment_Pass&""&vbcrlf
	Config=Config&"'�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Comment_IsPass:Sdcms_Comment_IsPass="&Sdcms_Comment_IsPass&""&vbcrlf
	Config=Config&"'Html��ǩ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadHtml:Sdcms_BadHtml="""&Sdcms_BadHtml&""""&vbcrlf
	Config=Config&"'Html��ǩ�¼�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadEvent:Sdcms_BadEvent="""&Sdcms_BadEvent&""""&vbcrlf
	Config=Config&"'�໰����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadText:Sdcms_BadText="""&Sdcms_BadText&""""&vbcrlf
	Config=Config&"'��Ϣ�����Զ���ȡ�ַ���,����Ϊ��200-500"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Length:Sdcms_Length=500"&vbcrlf
	Config=Config&"'���ݿ���������,TrueΪAccess��FalseΪMSSQL"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataType:Sdcms_DataType="&Sdcms_DataType&vbcrlf
	Config=Config&"'Access���ݿ�Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataFile:Sdcms_DataFile="""&Sdcms_DataFile&""""&vbcrlf
	Config=Config&"'Access���ݿ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataName:Sdcms_DataName="""&Sdcms_DataName&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ�IP,������ (local) �����IP "& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlHost:Sdcms_SqlHost="""&Sdcms_SqlHost&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlData:Sdcms_SqlData="""&Sdcms_SqlData&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ��ʻ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlUser:Sdcms_SqlUser="""&Sdcms_SqlUser&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlPass:Sdcms_SqlPass="""&Sdcms_SqlPass&""""&vbcrlf
	Config=Config&"'ϵͳ��װ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_CreateDate:Sdcms_CreateDate="""&Sdcms_CreateDate&""""&vbcrlf
	Config=Config&"%"
	Config=Config&">"&vbcrlf
	Config=Config&"<!--#include file=""../Skins/2009/Skins.asp""-->"
	Savefile "../Inc/","Const.asp",Config
	Echo Err.Description
	
	Progress 100

	Echo "<script>document.getElementById(""load_bnt"").value='�������';load_updates.innerHTML=""���ݿ�������ϣ��밴˵������������"";</script>"
	IF Err Then
		Echo "<script>document.getElementById(""load_bnt"").value='����ʧ��';</script>"
		Echo "<script>load_updates.style.width =""100%"";load_updates.innerHTML="""&Err.Description&""";</script>" & VbCrLf
		Err.Clear
	End IF
End Sub
%>
</body>
</html>
