<%
Dim Conn, ConnStr
Set Conn = Server.CreateObject("Adodb.Connection")
If IsDatabaseType = 1 Then
    ConnStr = "Provider = Sqloledb; User ID = " & DEF_SQLusername & "; Password = " & DEF_SQLpassword & "; Initial Catalog = " & DEF_SQLdatabase & "; Data Source = " & DEF_SQLserver & ";"
Else
    ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(""&SysRootDir&""&SiteDataPath&"/"&SiteDataAccess&"")
End If
Conn.Open ConnStr
If Err Then
    Err.Clear
    Set Conn = Nothing
    Response.Write ChinaQJLanguageTxt1
    Response.End
End If

'=======================================
'	SQL 查询 通用 函数 (Beta)
'	Conn	连接
'	sql		sql语句
'	返回一个字典集,如果没有返回空字典集
'======================================
Function SQL_Query(Conn,sql)
	Dim DataDiction, rs
	Set DataDiction=Server.CreateObject("Scripting.Dictionary")
	Set rs = server.createobject("adodb.recordset")
	rs.open sql,Conn,0,1
	If rs.eof Then 
		
	Else
		Dim dai
			dai=0
		Do While Not rs.eof
			c=rs.Fields.count
			Dim data
			Set data=Server.CreateObject("Scripting.Dictionary")
			For i=0 To (c-1)
				data.Add ""&rs.Fields(i).Name&"",""&rs(rs.Fields(i).Name)&""
			Next
			DataDiction.Add ""&dai&"",data
			dai=dai+1
			Set data=Nothing
			rs.movenext
		Loop
	End If	
	rs.close
	Set rs=Nothing
	Set SQL_Query=DataDiction
End Function

'==========================================
'	防SQL注入[简单版]
'	sql		语句
'	返回	替换结果
'===========================================
Function V(sql)
	V=replace(sql,"'","''")
End Function

'=============================================
'	SQL 操作通用函数[添加,修改,删除]
'	Conn		连接
'	sql			sql 语句
'	返回结果,受影响的行数
'===============================================
Function SQL_Other(Conn,sql)
	Dim rs
	Set rs = server.createobject("adodb.recordset")
	rs.open sql,Conn,1,3
	rs.close
	Set rs=Nothing
	Set SQL_Query=1
End Function

%>
<!--#include file="Function.Asp" -->