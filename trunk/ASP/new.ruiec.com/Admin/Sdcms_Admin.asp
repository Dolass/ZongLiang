<!--#include file="sdcms_check.asp"-->
<%
Dim Sdcms,title,Sd_Table,Sd_Table01,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set sdcms=New Sdcms_Admin
sdcms.Check_admin
Select Case action
	Case "add":title="����ʻ�"
	Case "edit":title="�޸��ʻ�"
	Case else:title="�ʻ�����"
End Select
Sd_Table="sd_admin"
Sd_Table01="sd_class"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="?action=add">����ʻ�</a>������<a href="?">�ʻ�����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li>
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 3:add
	Case "edit":add
	Case "save":save
	Case "admin":sdcms.Check_lever 4:admin
	Case "adminsave":adminsave
	Case "del":sdcms.Check_lever 5:del
	Case Else:sdcms.Check_lever 4:main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing
Sub Main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <tr>
      <td width="60" class="title_bg">���</td>
      <td class="title_bg">�ʻ�</td>
      <td width="80" class="title_bg">��¼����</td>
      <td width="120" class="title_bg">���һ��IP</td>
      <td width="80" class="title_bg">���</td>
      <td width="140" class="title_bg">����</td>
    </tr>
	<%
	Dim Sql,Rs
	Sql="Select id,sdcms_name,logintimes,lastip,isadmin From "&Sd_Table&" Order by ID Desc"
	Set Rs=Conn.Execute(sql)
	DbQuery=DbQuery+1
	While Not Rs.Eof
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
		<td align="center" height="25"><%=rs(0)%></td>
		<td><%=rs(1)%></td>
		<td align="center"><%=rs(2)%></td>
		<td align="center"><%=rs(3)%></td>
		<td align="center"><%=IIF(rs(4)=0,"��ͨ����Ա","��������Ա")%></td>
		<td align="center"><a href="?action=admin&id=<%=rs(0)%>">Ȩ�ޱ༭</a>��<a href="?action=edit&id=<%=rs(0)%>">�༭</a>��<a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%Rs.MoveNext:Wend:Rs.Close:Set Rs=Nothing%>
  </table>
<%
End Sub

Sub Add
	Dim Rs,ID,t0,t1,t2
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	Check_Add
	IF ID>0 Then
		Set Rs=Conn.Execute("Select sdcms_name,penname,isadmin from "&Sd_Table&" where ID="&ID&"")
		DbQuery=DbQuery+1
		IF Rs.Eof then
			Echo "��������":Exit Sub
		Else
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
		End IF
		Rs.Close
		Set Rs=Nothing
	Else
		t2=0
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
    <tr>
      <td width="120" align="center" class="tdbg">�� ����      </td>
      <td class="tdbg"><%IF ID=0 Then%><input name="t0" type="text" class="input" size="30" maxlength="20" onkeyup="$('#t2')[0].value=this.value"><%End IF%><%=t0%></td>
    </tr>
    <tr class="tdbg">
      <td align="center">�� �룺      </td>
      <td><input name="t1" type="text" class="input" size="30"  maxlength="20"><%=IIF(ID=0,"��<span>���޸�������</span>","")%></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�� ����      </td>
      <td><input name="t2" type="text" class="input" size="30" value="<%=t1%>" id="t2" maxlength="20">��<span>���ڷ�����Ϣʱ��������</span></td>
    </tr>
	<%IF Load_Cookies("sdcms_admin")=1 then%>
    <tr class="tdbg">
      <td align="center">�� ��</td>
      <td><input name="t3" type="radio" value="0" <%=IIF(t2=0,"checked","")%> id="t3_0"><label for="t3_0">��ͨ����Ա</label> <input name="t3" type="radio"  value="1" <%=IIF(t2=1,"checked","")%> id="t3_1"><label for="t3_1">��������Ա</label></td>
    </tr>
	<%End IF%>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="�� ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Save
	Dim t0,t1,t2,t3,Rs,Sql,LogMsg,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=FilterText(Trim(Request.Form("t0")),2)
	t1=FilterText(Trim(Request.Form("t1")),2)
	t2=FilterText(Trim(Request.Form("t2")),2)
	t3=IsNum(Trim(Request.Form("t3")),0)
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select sdcms_name,sdcms_pwd,penname,isadmin,logintimes From "&Sd_Table&" where "
	IF ID=0 Then 
		Sql=Sql&"sdcms_name='"&t0&"'"
	Else
		Sql=Sql&"id="&ID
	End IF

	Rs.Open Sql,Conn,1,3
	IF ID=0 Then
	  IF Not Rs.Eof Then
		  Echo "���û����Ѵ��ڣ��뻻������":Died
	  End IF
	End IF
	IF ID=0 Then 
	  Rs.Addnew
	Else
	  Rs.Update
	End IF
	IF ID=0 Then 
		Rs(0)=Left(t0,20)
	End IF
	IF ID=0 Then
		Rs(1)=Md5(t1)
	Else
	  IF Len(t1)>0 Then
		  Rs(1)=md5(t1)
		  IF Clng(sdcms_adminid)=Clng(ID) Then Add_Cookies "sdcms_pwd",md5(Rs(1))
	  End IF
	End IF
	Rs(2)=Left(t2,20)
	IF Load_Cookies("sdcms_admin")=1 then
		Rs(3)=t3
	End IF
	
	IF ID=0 Then LogMsg="��ӹ����ʻ�" Else LogMsg="�޸Ĺ����ʻ�":Add_Cookies "sdcms_all_lever",t3
	AddLog sdcms_adminname,GetIp,LogMsg&rs(0),0
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	IF Load_Cookies("sdcms_admin")=1 Then
		Go("?")
	Else
		Alert "�����޸ĳɹ�","?id="&ID&"&action=edit"
	End IF
End Sub

Sub Admin
	Dim Rs,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	Check_Add
	Set Rs=Conn.Execute("select id,sdcms_name,isadmin,alllever,infolever from "&Sd_Table&" where id="&id&"")
	DbQuery=DbQuery+1
	IF Rs.Eof Then
		Echo "����Ƿ��ύ����":Exit Sub
	ElseIF Rs(2)=1 Then
		Echo "��������Ա����༭Ȩ��":Exit Sub
	End IF
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
  <form method="post" action="?action=adminsave&id=<%=id%>">
    <tr>
      <td colspan="2" align="center" class="tdbg01">����Ա��<%=Rs(1)%> ��Ȩ������</td>
      </tr>
    <tr class="tdbg">
      <td width="120" height="25" align="center">�١�������</td>
      <td><input name="chkAll" type="checkbox" id="chkAll" Onclick="CheckAll(this.form)" value="checkbox"><label for="chkall">ȫѡ</label></td>
    </tr>
    <tr class="tdbg">
      <td align="center">ȫ��Ȩ�ޣ�</td>
      <td>
	  <%
	  Dim Menu(5,15)
	  Menu(0,0)="ϵͳ����"
	  Menu(0,1)="ϵͳ����|1"
	  Menu(0,2)="��־����|2"
	  Menu(0,3)="ϵͳ�ʻ�(���)|3"
	  Menu(0,4)="ϵͳ�ʻ�(�༭)|4"
	  Menu(0,5)="ϵͳ�ʻ�(ɾ��)|5"
	  
	  Menu(1,0)="��Ϣ����"
	  Menu(1,1)="������(���)|6"
	  Menu(1,2)="������(�༭)|7"
	  Menu(1,3)="������(ɾ��)|8"
	  Menu(1,4)="ר�����(���)|9"
	  Menu(1,5)="ר�����(�༭)|10"
	  Menu(1,6)="ר�����(ɾ��)|11"
	  Menu(1,7)="��Ϣ����(���)|12"
	  Menu(1,8)="��Ϣ����(�༭)|13"
	  Menu(1,9)="��Ϣ����(ɾ��)|14"
	  Menu(1,10)="��ҳ����(���)|15"
	  Menu(1,11)="��ҳ����(�༭)|16"
	  Menu(1,12)="��ҳ����(ɾ��)|17"
	  
	  Menu(2,0)="���ӹ���"
	  Menu(2,1)="���|18"
	  Menu(2,2)="�༭|19"
	  Menu(2,3)="ɾ��|20"
	 
	  
	  Menu(3,0)="�������"
	  Menu(3,1)="���|21"
	  Menu(3,2)="�༭|22"
	  Menu(3,3)="ɾ��|23"
	  Menu(3,4)="����|24"
	  
	  Menu(4,0)="�������"
	  Menu(4,1)="��Ƭ����(���)|25"
	  Menu(4,2)="��Ƭ����(�༭)|26"
	  Menu(4,3)="��Ƭ����(ɾ��)|27"
	  Menu(4,4)="ģ�����(���)|28"
	  Menu(4,5)="ģ�����(�༭)|29"
	  Menu(4,6)="ģ�����(ɾ��)|30"
	
	  
	  Menu(5,0)="���ɹ���"
	  Menu(5,1)="���ɹ���|31"

	  Dim I,J,t0
	  For I=0 To Ubound(Menu,1)
		  Echo "<b>"&Menu(I,0)&"</b><br>"
		  For J=1 To Ubound(Menu,2)
			  IF IsEmpty(menu(I,J)) Then exit for
			  t0=Split(Menu(I,J),"|")
			  Echo "<div style=""float:left;width:19%;margin:2px 0;""><input Type=""checkbox"" value="""&t0(1)&""" id=""admin_"&t0(1)&""" name=""t0"""
			  IF Instr(", "&rs(3)&", ",", "&t0(1)&", ")>0 Then Echo "checked" End IF
			  Echo "><label for=""admin_"&t0(1)&""">"&t0(1)&".<span>"&t0(0)&"</label></span></div>"
		  Next
		  Echo "<div class=""clear mag_t""></div>"
	  Next
	  %>	  </td>
    </tr>
	<tr class="tdbg">
      <td align="center" height="25">��ĿȨ�ޣ�</td>
      <td>˵������ĿȨ�޾��м̳��ԣ��ұ�����ȫ��Ȩ���о�����Ӧ����Ȩ�ޡ�<div style="overflow:auto;width:99%;height:100px;"><%Class_Lever Rs(4)%></div></td>
	</tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="�� ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
	Rs.Close
	Set Rs=Nothing
End Sub

Sub AdminSave
	Dim t0,t1,rs,sql,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=FilterHtml(Trim(Request.Form("t0")))
	t1=FilterHtml(Trim(Request.Form("t1")))
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select Alllever,InfoLever From "&Sd_Table&" Where Id="&Id&""
	Rs.Open Sql,Conn,1,3
	Rs.Update
	Rs(0)=t0
	Rs(1)=t1
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Alert "Ȩ�ޱ���ɹ�","?"
End Sub

Sub Del
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF Clng(sdcms_adminid)=Clng(ID) Then
		Alert "����ɾ���Լ�","?":Exit Sub
	Else
		Dim LogMsg
		LogMsg="ɾ�������ʻ���"
		AddLog sdcms_adminname,GetIp,LogMsg&loadrecord("sdcms_name",Sd_Table,id),0
		Conn.Execute("Delete From "&Sd_Table&" Where Id="&id&"")
		Go("?")
	End IF
End Sub

Sub Class_Lever(ByVal t0)
	Dim t1:t1=Get_Class_Array
	IF IsArray(t1) Then
		Lever_Class 0,t0,t1
		Echo LeverClass
	End IF
End Sub

Dim LeverClass
Sub Lever_Class(ByVal t0,ByVal t1,t2)
	Dim Class_Array,I,J,Rows
	Class_Array=t2
	Rows=UBound(Class_Array,2)
	For I=0 To Rows
		IF Class_Array(3,I)=t0 Then
			LeverClass=LeverClass&"<input type=""checkbox"" name=""t1"" value="""&Class_Array(0,I)&"|1"" id=""class_"&Class_Array(0,I)&"|1"""
			IF Instr(", "&t1&", ",", "&Class_Array(0,I)&"|1"&", ")>0 Then LeverClass=LeverClass& " checked" End IF
			LeverClass=LeverClass&"> <span><label for=""class_"&Class_Array(0,I)&"|1"">���</label></span>"
			LeverClass=LeverClass&"<input type=""checkbox"" name=""t1"" value="""&Class_Array(0,I)&"|2"" id=""class_"&Class_Array(0,I)&"|2"""
			IF Instr(", "&t1&", ",", "&Class_Array(0,I)&"|2"&", ")>0 Then LeverClass=LeverClass& " checked" End IF
			LeverClass=LeverClass&"> <span><label for=""class_"&Class_Array(0,I)&"|2"">�༭</label></span>"
			LeverClass=LeverClass&"<input type=""checkbox"" name=""t1"" value="""&Class_Array(0,I)&"|3"" id=""class_"&Class_Array(0,I)&"|3"""
			IF Instr(", "&t1&", ",", "&Class_Array(0,I)&"|3"&", ")>0 Then LeverClass=LeverClass& " checked" End IF
			LeverClass=LeverClass&"> <span><label for=""class_"&Class_Array(0,I)&"|3"">ɾ��</label></span>��"
			For J=0 To Class_Array(2,I)-1
				LeverClass=LeverClass&"��"
			Next
			LeverClass=LeverClass&IIF(Class_Array(3,I)>0,"��","")&Class_Array(1,I)&"<br>"
			Lever_Class Class_Array(0,I),t1,t2
		End IF
	Next
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	IF action="add" Then
		Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
		Check_Add=Check_Add&"	{"&vbcrlf
		Check_Add=Check_Add&"	alert('�ʻ����Ʋ���Ϊ��');"&vbcrlf
		Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
		Check_Add=Check_Add&"	return false"&vbcrlf
		Check_Add=Check_Add&"	}"&vbcrlf
		Check_Add=Check_Add&"	if (document.add.t1.value=='')"&vbcrlf
		Check_Add=Check_Add&"	{"&vbcrlf
		Check_Add=Check_Add&"	alert('�ʻ����벻��Ϊ��');"&vbcrlf
		Check_Add=Check_Add&"	document.add.t1.focus();"&vbcrlf
		Check_Add=Check_Add&"	return false"&vbcrlf
		Check_Add=Check_Add&"	}"&vbcrlf
	End IF
	Check_Add=Check_Add&"	if (document.add.t2.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��������Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t2.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>