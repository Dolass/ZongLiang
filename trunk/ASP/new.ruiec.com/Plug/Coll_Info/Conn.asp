<%
'����Ϊ�ɼ����ݿ��������ã�������Լ�����Ҫ�޸�
Dim Coll_Conn
Sub Collection_Data
	Dim Collection_DbFile,Collection_DbName,Connstr

	'���Ŀ¼ͬ������
	Collection_DbFile=Sdcms_DataFile
	Collection_DbName="Collection.Mdb"

	Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(Sdcms_root&Collection_DbFile&"/"&Collection_DbName)
	
	On Error Resume Next
	Set Coll_Conn=Server.CreateObject("ADODB.Connection")
	Coll_Conn.Open Connstr
	IF Err Then
		Err.Clear
		Set Coll_Conn=Nothing
		Echo "�ɼ����ݿ����ӳ�������Plug/Coll_Info/Conn.asp�����ļ�"
		Died
	End If
End Sub
%>