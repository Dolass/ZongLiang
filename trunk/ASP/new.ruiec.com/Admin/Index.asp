<!--#include file="sdcms_check.asp"-->
<%
Const loginnum=3 '��¼ʧ�ܺ󣬽�ֹ�ĵ�¼�Ĵ������������Ϊ�ܴ����֣�
Sub Main()
	IF Len(Load_Cookies("sdcms_name"))>0 And Len(Load_Cookies("sdcms_pwd"))>0 Then
		Go("sdcms_index.asp"):Exit Sub
	End IF
	Echo"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
	Echo"<html xmlns=""http://www.w3.org/1999/xhtml"">"
	Echo"<head>"
	Echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"
	Echo"<title>��վ��Ϣ����ϵͳ Powered By Sdcms</title>"
	Echo"<link href=""Css/home.css"" rel=""stylesheet"" type=""text/css"" />"
	Echo"</head>"
	Echo"<body>"
	Echo"<Script>if(self!=top){top.location=self.location;}</script>"
	Echo"<div class=""b_w"">"
	Echo"   <div class=""l_t"">"
	Echo"      <div class=""left l_title"">��վ����ϵͳ</div>"
	Echo"	  <div class=""right""><a href=""../""><img src=""images/icon_back.gif"" alt=""����"" hspace=""4"" vspace=""8"" border=""0""  /></a><a href=""javascript:window.close()""><img src=""images/icon_close.gif"" alt=""�ر�"" hspace=""4"" vspace=""8"" border=""0"" /></a></div>"
	Echo"   </div>"
	Echo"  <div class=""l_bg"">"
	Echo"    <ul class=""l_user"">"
	Echo"<script language=""javascript"">"
	Echo"function checklogin()"
	Echo"{"
	Echo"  if(document.login.username.value=='')"
	Echo"     {alert('�������ʻ�');"
	Echo"      document.login.username.focus();"
	Echo"      return false"
	Echo"    }"
	Echo"  if (document.login.password.value=='')"
	Echo"   {alert('����������');"
	Echo"    document.login.password.focus();"
	Echo"    return false"
	Echo"   }"
	Echo"   if (document.login.yzm.value=='')"
	Echo"   {alert('��������֤��');"
	Echo"    document.login.yzm.focus();"
	Echo"    return false"
	Echo"   }"
	Echo"}"
	Echo"</script>"
	Echo"<form action=""index.asp?action=check"" name=""login"" method=""post"" onSubmit=""return checklogin();"">"
	Echo"	  <li>�ʻ���<input name=""username"" size=""14"" type=""text"" class=""l_input"" /></li>"
	Echo"	  <li>���룺<input name=""password"" size=""14"" type=""password"" class=""l_input"" /></li>"
	Echo"	  <li>��֤��<input name=""yzm"" size=""3"" type=""text"" class=""l_input"" /> <img src=""../inc/sdcmscode.asp?t0=60&t1=14"" align=""absmiddle"" title=""�������?����ˢ����֤��"" style=cursor:pointer onClick=""this.src+='&'+Math.random();""></li>"
	Echo"	  <li><input class=""l_bnt"" value=""�� ¼"" type=""submit"" />��<input class=""l_bnt"" value=""�� д"" type=""reset"" /></li>"
	Echo"	  </form>"
	Echo"	</ul>"
	Echo"  </div>"
	Echo"  <div class=""l_f"">"
	Echo"    <div class=""left""><img src=""images/f_l.gif"" /></div>"
	Echo"	<div class=""left""><img src=""images/f_bg.gif"" width=""378"" height=""36"" /></div>"
	Echo"	<div class=""right""><img src=""images/f_r.gif"" /></div>"
	Echo"  </div>"
	Echo"</div>"
	
	Echo"</body>"
	Echo"</html>"
End Sub

Sub Check
	Dim username,password,code,getcode,Rs
	IF Check_post Then Echo "1��ֹ���ⲿ�ύ����!":Exit Sub
	username=FilterText(Trim(Request.Form("username")),2)
	password=FilterText(Trim(Request.Form("password")),2)
	code=Trim(Request.Form("yzm"))
	getcode=Session("SDCMSCode")
	IF errnum>=loginnum Then Echo "ϵͳ�ѽ�ֹ�������ٵ�¼":died
	IF code="" Then Alert "��֤�벻��Ϊ�գ�","javascript:history.go(-1)":Died
	IF code<>"" And Not Isnumeric(code) Then Alert "��֤�����Ϊ���֣�","javascript:history.go(-1)":Died
	IF code<>getcode Then Alert "��֤�����","javascript:history.go(-1)":Died
	IF username="" or password="" Then
		Echo "�û��������벻��Ϊ��":Died
	Else
		Set Rs=Conn.Execute("Select Id,Sdcms_Name,Sdcms_Pwd,isadmin,alllever,infolever From Sd_Admin Where Sdcms_name='"&username&"' And Sdcms_Pwd='"&md5(password)&"'")
		IF Rs.Eof Then
			AddLog username,GetIp,"��¼ʧ��",1
			Echo "�û������������,���ջ��� "&loginnum-errnum&" �λ���"
		Else
			Add_Cookies "sdcms_id",Rs(0)
			Add_Cookies "sdcms_name",username
			Add_Cookies "sdcms_pwd",Rs(2)
			Add_Cookies "sdcms_admin",Rs(3)
			Add_Cookies "sdcms_alllever",Rs(4)
			Add_Cookies "sdcms_infolever",Rs(5)
			Conn.Execute("Update Sd_Admin Set logintimes=logintimes+1,LastIp='"&GetIp&"' Where id="&Rs(0)&"")
			AddLog username,GetIp,"��¼�ɹ�",1
			'�Զ�ɾ��30��ǰ��Log��¼
			IF Sdcms_DataType Then
				Conn.Execute("Delete From Sd_Log Where DateDiff('d',adddate,Now())>30")
			Else
				Conn.Execute("Delete From Sd_Log Where DateDiff(d,adddate,GetDate())>30")
			End IF
			Go("sdcms_index.asp")
		End IF
		Rs.Close
		Set Rs=Nothing
	End IF
End Sub

Sub Out
	AddLog sdcms_adminname,GetIp,"�˳���¼",1
	Add_Cookies "sdcms_id",Empty
	Add_Cookies "sdcms_name",Empty
	Add_Cookies "sdcms_pwd",Empty
	Add_Cookies "sdcms_admin",Empty
	Add_Cookies "sdcms_alllever",Empty
	Add_Cookies "sdcms_infolever",Empty
	Go "?"
End Sub

Function ErrNum
	Dim Sql
	DbOpen
	Sql="Select Count(ID) From Sd_Log Where Ip='"&GetIp&"' And Content Like '��¼ʧ��' And "
	IF Sdcms_DataType Then Sql=Sql&" Adddate>=Date()" Else Sql=Sql&" Adddate>=Getdate()"
	Errnum=Conn.Execute(Sql)(0)
End Function

Dim Action:Action=Lcase(Trim(Request.QueryString("Action")))
Select Case Action
	Case "check":Check
	Case "out":Out
	Case Else:main
End Select
Closedb
%>