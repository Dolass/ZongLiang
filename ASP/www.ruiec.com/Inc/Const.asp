<%
'��վ����
    Dim Sdcms_WebName:Sdcms_WebName="��ҵչ��,����Ӫ��,���繫��,Ʒ�Ʋ߻�,����Ӫ��-����ý"
'��վ����
    Dim Sdcms_WebUrl:Sdcms_WebUrl="http://127.0.0.1"
'ϵͳĿ¼����Ŀ¼Ϊ��"/",����Ŀ¼��ʽΪ��"/sdcms/"
    Dim Sdcms_Root:Sdcms_Root="/"
'����ģʽ,0Ϊ��̬Ĭ�ϣ�1Ϊα��̬��2Ϊ��̬
    Dim Sdcms_Mode:Sdcms_Mode=2
'����ģʽ
    Dim blogmode:blogmode=false
'����Ŀ¼����Ŀ¼Ϊ��,Ҳ����ָ��ΪĳһĿ¼����ʽΪ��"sdcms/",�����ԣ�"/"����
    Dim Sdcms_HtmDir:Sdcms_HtmDir=""
'ʱ������,�벻Ҫ�������
    Dim Sdcms_TimeZone:Sdcms_TimeZone=0
'�Ƿ���������־
    Dim Sdcms_AdminLog:Sdcms_AdminLog=True
'�Ƿ�������
    Dim Sdcms_Cache:Sdcms_Cache=False
'����ǰ׺,����ж��վ�㣬�����ò�ͬ��ֵ
    Dim Sdcms_Cookies:Sdcms_Cookies="7Hh3Qb4Xf1Gl"
'����ʱ�䣬��λΪ��
    Dim Sdcms_CacheDate:Sdcms_CacheDate=60
'ϵͳ�ļ��ĺ�׺�������鲻Ҫ�Ķ�
    Dim Sdcms_FileTxt:Sdcms_FileTxt=".html"
'ϵͳģ��Ŀ¼,һ�㲻�����ֶ�����
    Dim Sdcms_Skins_Root:Sdcms_Skins_Root="2009"
'ϵͳ�����ļ���Ĭ���ļ��������鲻Ҫ�Ķ�
    Dim Sdcms_FileName:Sdcms_FileName="{ID}"
'����Ŀ¼
    Dim Sdcms_UpfileDir:Sdcms_UpfileDir="Upfile"
'�����ϴ��ļ�����,�������|��
    Dim Sdcms_UpfileType:Sdcms_UpfileType="jpe|gif|jpg|png|bmp|rar|zip|flv|swf"
'�����ϴ����ļ����ֵ����λ��KB
    Dim Sdcms_upfileMaxSize:Sdcms_upfileMaxSize=200
'����ƫ������,ϵͳ�Զ����ɣ���������Ķ�
    Dim Sdcms_Create_Set:Sdcms_Create_Set="0, 1, 2"
'����GOOGLE��ͼ�Ĳ���
    Dim Sdcms_Create_GoogleMap:Sdcms_Create_GoogleMap=Split("20,daily,0.8",",")
'���ۿ���
    Dim Sdcms_Comment_Pass:Sdcms_Comment_Pass=1
'�������
    Dim Sdcms_Comment_IsPass:Sdcms_Comment_IsPass=1
'Html��ǩ����
    Dim Sdcms_BadHtml:Sdcms_BadHtml="Table|TBODY|TR|TD|Body|Meta|iframe|SCRIPT|form|style|div|object|TEXTAREA"
'Html��ǩ�¼�����
    Dim Sdcms_BadEvent:Sdcms_BadEvent="javascript:|Document.|onerror|onload|onmouseover"
'�໰����
    Dim Sdcms_BadText:Sdcms_BadText="������|Fuck|TMD|NND|�Ҳ�|����|���"
'��Ϣ�����Զ���ȡ�ַ���,����Ϊ��200-500
    Dim Sdcms_Length:Sdcms_Length=500
'���ݿ���������,TrueΪAccess��FalseΪMSSQL
    Dim Sdcms_DataType:Sdcms_DataType=True
'Access���ݿ�Ŀ¼
    Dim Sdcms_DataFile:Sdcms_DataFile="Data"
'Access���ݿ�����
    Dim Sdcms_DataName:Sdcms_DataName="#%@ruie.com_2012@%#.mdb"
'MSSQL���ݿ�IP,������ (local) �����IP 
    Dim Sdcms_SqlHost:Sdcms_SqlHost="(local)"
'MSSQL���ݿ���
    Dim Sdcms_SqlData:Sdcms_SqlData="ruiec2012"
'MSSQL���ݿ��ʻ�
    Dim Sdcms_SqlUser:Sdcms_SqlUser="ruiec2012"
'MSSQL���ݿ�����
    Dim Sdcms_SqlPass:Sdcms_SqlPass="admin888"
'ϵͳ��װ����
    Dim Sdcms_CreateDate:Sdcms_CreateDate="2011-1-14 7:52:45"
%>
<!--#include file="../Skins/2009/Skins.asp"-->