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
<title>SDCMS������������</title>
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
		<h1>SDCMS������������</h1>
		<dl>
			<dt class="update_content"><strong>����˵����</strong>
				<ul>
					<li class="action">��������SDCMS1.3 To SDCMS1.3.1��</li>
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
dim action
action=request("action")
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
'���ã�1.3 To 1.3.1
'Author��ITƽ��
'Date��2011-10-10
'=========================================
	Echo "<script>document.getElementById(""load_bnt"").disabled=true;</script>"
	
	Progress 5
	Progress 10
	Progress 5
	Progress 20
	Progress 25
	Progress 30
	Progress 35
	Progress 40
	Progress 45
	Progress 50
	Progress 55
	Progress 60

	if Sdcms_Mode=1 then
	dim httpd
		httpd=""
		httpd=httpd&"[ISAPI_Rewrite]"&vbcrlf&vbcrlf

		httpd=httpd&"# 3600 = 1 hour"&vbcrlf
		httpd=httpd&"#CacheClockRate 3600"&vbcrlf&vbcrlf
		
		httpd=httpd&"RepeatLimit 32"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д��ҳ"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Page/(.*)_(\d*)\.html "&Get_Sys_Dir&"page/\?id=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Page/(.*)\.html "&Get_Sys_Dir&"page/\?id=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дר��"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Topic/List_(.*)_(\d*)\.html "&Get_Sys_Dir&"Topic/\List.asp\?ID=$1&Page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Topic/List_(.*)\.html "&Get_Sys_Dir&"Topic/\List.asp\?ID=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Topic/Index_(.*)\.html "&Get_Sys_Dir&"Topic/\?page=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"Topic/Index\.html "&Get_Sys_Dir&"Topic/\index.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дTags"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"tags/List_(.*)\.html "&Get_Sys_Dir&"tags/\List.asp\?page=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"tags/List\.html(.*) "&Get_Sys_Dir&"tags/\list.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"tags/(.*)_(\d*)\.html "&Get_Sys_Dir&"tags/\?/$1/$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"tags/(.*)\.html "&Get_Sys_Dir&"tags/\?/$1/ [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дindex.asp"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"index\.html(.*) "&Get_Sys_Dir&"index.asp [N,I]"&vbcrlf&vbcrlf
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"index_(\d*)\.html "&Get_Sys_Dir&"index.asp\?page=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дsitemap.asp"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"sitemap\.html(.*) "&Get_Sys_Dir&"sitemap.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"html/(.*)/(\d*)_(\d*)\.html "&Get_Sys_Dir&"Info/\View.asp\?ID=$2&page=$3 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"html/(.*)/(\d*)\.html "&Get_Sys_Dir&"Info/\View.asp\?ID=$2 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д�б�"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"html/(.*)/(\d*) "&Get_Sys_Dir&"Info/\?cname=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"html/(.*)/ "&Get_Sys_Dir&"Info/\?cname=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"date-(.*)_(\d*) "&Get_Sys_Dir&"date.asp\?c=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Get_Sys_Dir&"date-(.*) "&Get_Sys_Dir&"date.asp\?c=$1 [N,I]"
		Savefile "/","httpd.ini",httpd
	end if
	Progress 65
	Progress 70
	Progress 75
	Progress 80
	'������ϵͳ����
	dim Config
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
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Mode:Sdcms_Mode="&Sdcms_Mode&""&vbcrlf
	Config=Config&"'����ģʽ"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim blogmode:blogmode=false"&vbcrlf
	Config=Config&"'����Ŀ¼����Ŀ¼Ϊ��,Ҳ����ָ��ΪĳһĿ¼����ʽΪ��""sdcms/"",�����ԣ�""/""����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_HtmDir:Sdcms_HtmDir="""&Sdcms_HtmDir&""""&vbcrlf
	Config=Config&"'ʱ������,�벻Ҫ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_TimeZone:Sdcms_TimeZone="&Sdcms_TimeZone&""&vbcrlf
	Config=Config&"'�Ƿ���������־"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_AdminLog:Sdcms_AdminLog="&Sdcms_AdminLog&""&vbcrlf
	Config=Config&"'�Ƿ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cache:Sdcms_Cache="&Sdcms_Cache&""&vbcrlf
	Config=Config&"'����ǰ׺,����ж��վ�㣬�����ò�ͬ��ֵ"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cookies:Sdcms_Cookies="""&Sdcms_Cookies&""""&vbcrlf
	Config=Config&"'����ʱ�䣬��λΪ��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_CacheDate:Sdcms_CacheDate="&Sdcms_CacheDate&""&vbcrlf
	Config=Config&"'ϵͳ�ļ��ĺ�׺�������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileTxt:Sdcms_FileTxt="""&Sdcms_FileTxt&""""&vbcrlf
	Config=Config&"'ϵͳģ��Ŀ¼,һ�㲻�����ֶ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Skins_Root:Sdcms_Skins_Root="""&Sdcms_Skins_Root&""""&vbcrlf
	Config=Config&"'ϵͳ�����ļ���Ĭ���ļ��������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileName:Sdcms_FileName="""&Sdcms_FileName&""""&vbcrlf
	Config=Config&"'����Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileDir:Sdcms_UpfileDir="""&Sdcms_UpfileDir&""""&vbcrlf
	Config=Config&"'�����ϴ��ļ�����,�������|��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileType:Sdcms_UpfileType="""&Sdcms_UpfileType&""""&vbcrlf
	Config=Config&"'�����ϴ����ļ����ֵ����λ��KB"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_upfileMaxSize:Sdcms_upfileMaxSize="&Sdcms_upfileMaxSize&""&vbcrlf
	Config=Config&"'����ƫ������,ϵͳ�Զ����ɣ���������Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Create_Set:Sdcms_Create_Set="""&Sdcms_Create_Set&""""&vbcrlf
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
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Length:Sdcms_Length="&Sdcms_Length&""&vbcrlf
	Config=Config&"'���ݿ���������,TrueΪAccess��FalseΪMSSQL"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataType:Sdcms_DataType="&Sdcms_DataType&vbcrlf
	Config=Config&"'Access���ݿ�Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataFile:Sdcms_DataFile=""Data"""&vbcrlf
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
	Config=Config&"<!--#include file=""../Skins/"&Sdcms_Skins_Root&"/Skins.asp""-->"
	Savefile "../Inc/","Const.asp",Config
	Progress 85
	Progress 90
	Progress 95
	Progress 96
	Progress 97
	Progress 98
	Progress 99
	Progress 100

	Echo "<script>document.getElementById(""load_bnt"").value='�������';load_updates.innerHTML=""������ϣ��밴˵������������"";</script>"
	
End Sub
%>
</body>
</html>
