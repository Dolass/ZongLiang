<!--#include file="sdcms_check.asp"-->
<%
Dim Sdcms,title,Sd_Table,ordnum,stype,t,Action
Action=Lcase(Trim(Request.QueryString("Action")))
t=Lcase(Trim(Request.QueryString("t")))
Set sdcms=New Sdcms_Admin
sdcms.Check_admin
Select Case action
	Case "add":title="�������"
	Case "edit":title="�޸�����"
	Case Else:title="���ӹ���"
End Select
Sd_Table="sd_link"
ordnum=trim(request("ordnum"))
stype=trim(request("stype"))
stype=IsNum(stype,0)
Sdcms_Head
%>

<div class="sdcms_notice"><span>���������</span><a href="?action=add">�������</a>������<a href="?">���ӹ���</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="<%IF stype<>0 Then Echo "un" End IF%>sub"><a<%IF stype<>"" Then%> href="?"<%End IF%>><%=title%></a></li>
	<%IF stype<>"" Then%><li class="<%IF stype<>1 Then Echo "un" End IF%>sub"><a href="?stype=1">ͼƬ����</a></li><li class="<%IF stype<>2 Then Echo "un" End IF%>sub"><a href="?stype=2">��������</a></li><li class="<%IF stype<>3 Then Echo "un" End IF%>sub"><a href="?stype=3">�������</a></li><%End IF%>
</ul>
<div id="sdcms_right_b">
<%
Select Case action
	Case "add":sdcms.Check_lever 21:add
	Case "edit":sdcms.Check_lever 22:add
	Case "save":save
	Case "del":sdcms.Check_lever 23:del
	Case "up":sdcms.Check_lever 22:up
	Case "down":sdcms.Check_lever 22:down
	Case "pass":sdcms.Check_lever 22:pass(1)
	Case "nopass":sdcms.Check_lever 22:pass(0)
	Case Else:main
End Select
Db_Run
closedb
Set sdcms=Nothing
Sub main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <tr>
      <td width="60" class="title_bg">���</td>
      <td class="title_bg">��վ����</td>
      <td width="60" class="title_bg">״̬</td>
      <td width="60" class="title_bg">����</td>
      <td width="120" class="title_bg">��ʶ</td>
      <td width="120" class="title_bg">����</td>
    </tr>
	<%
	Dim Sql,Rs
	Sql="select id,title,url,pic,ispic,ordnum,ispass,content from "&Sd_Table&" where "
	Select Case stype
		Case "1":Sql=sql&"ispic=1 and ispass=1"
		Case "2":Sql=sql&"ispic=0 and ispass=1"
		Case "3":Sql=sql&"ispass=0"
	Case Else:Sql=sql&"ispass=1"
	End Select
	Sql=Sql&" order by ordnum desc"
	Set Rs=Conn.Execute(sql)
	DbQuery=DbQuery+1
	While Not Rs.Eof
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td align="center"><%=rs(0)%></td>
	  <td height="33" align="center"><%IF rs(4)=1 Then%><a href="<%=rs(2)%>"><img src="<%=rs(3)%>" border="0" width="88" height="31"></a><%Else%><a href="<%=rs(2)%>"><%=rs(1)%></a><%End IF%></td>
	  <td align="center"><%=IIF(rs(6)=0,"δ��֤","����֤")%></td>
	  <td align="center"><a href="?action=up&id=<%=rs(0)%>&ordnum=<%=rs(5)%>&stype=<%=stype%>&t0=<%=rs(4)%>">��</a> <a href="?action=down&id=<%=rs(0)%>&ordnum=<%=rs(5)%>&stype=<%=stype%>&t0=<%=rs(4)%>">��</a></td>
      <td align="center"><%=rs(7)%></td>
      <td align="center"><%IF stype=3 Then%><a href="?action=pass&id=<%=rs(0)%>&stype=<%=stype%>">ͨ����֤</a><%Else%><a href="?action=nopass&id=<%=rs(0)%>&stype=<%=stype%>">ȡ����֤</a><%End IF%> <a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>&stype=<%=stype%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%Rs.Movenext:Wend:Rs.Close:Set Rs=Nothing%>
  </table>
  
<%
End Sub

Sub Add
	Dim Rs
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID>0 Then
		Set Rs=Conn.Execute("select id,title,url,ispic,pic,ispass,content from "&Sd_Table&" where id="&id&"")
		IF Rs.Eof Then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3,t4,t5,t6
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
			t4=Rs(4)
			t5=Rs(5)
			t6=Rs(6)
		End IF
		Rs.Close
		Set Rs=Nothing
	Else
		t3=1
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
    <tr>
      <td width="120" align="center" class="tdbg">��վ���ƣ�      </td>
      <td class="tdbg"><input name="t0" type="text" class="input" value="<%=t1%>" id="t0" size="30"></td>
    </tr>
    <tr class="tdbg">
      <td align="center">��վ������      </td>
      <td><input name="t1" type="text" class="input" value="<%=t2%>" id="t1" size="30">��<span>����д����·�� �磺http://www.sdcms.cn</span></td>
    </tr>
    <tr class="tdbg">
      <td align="center">�������</td>
      <td><input name='t2' type='radio' onClick=$('#flag')[0].style.display='none';this.form.t3.disabled=true; value='0' <%=IIF(t3=0,"checked","")%> id="t2_0"><label for="t2_0">��������</label> <input name='t2' type='radio' onClick=$('#flag')[0].style.display='';this.form.t3.disabled=false; value='1' <%=IIF(t3=1,"checked","")%> id="t2_1"><label for="t2_1">ͼƬ����</label></td>
    </tr>
	  <tr class="tdbg <%IF t3=0 Then%>dis<%End IF%>" id='flag' >
      <td align="center">ͼƬ��ַ��</td>
      <td><input name="t3" type="text" value="<%=t4%>" class="input" id="t3" size="40"  <%IF t3=0 Then%>disabled<%End IF%>><%admin_upfile 1,"100%","20","t3","UpLoadIframe",0,0%></td>
    </tr>
	<tr class="tdbg">
      <td align="center">���ӱ�ʶ��</td>
      <td><input name="t5" type="text" value="<%=t6%>" class="input" id="t5" size="40" maxlength="50">��<span>����Ϊ�գ����������������͵ı�ʶ</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�������ԣ�      </td>
      <td><input name="t4" id="t4" type="checkbox" value="1" <%=IIF(t5=1,"checked","")%> /><label for="t4">����֤</label><%IF Sdcms_Mode=2 Then%>��<input name="up1" id="up1" type="checkbox" value="1" /><label for="up1">������ҳ</label><%End IF%></td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��������"> <input type="button" onClick="history.go(-1)" class="bnt" value="��������"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Save
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim t0,t1,t2,t3,t4,t5,t6,up1,Rs,Sql,LogMsg,sdcms_c
	t0=FilterText(Trim(Request.Form("t0")),1)
	t1=FilterText(Trim(Request.Form("t1")),0)
	t2=IsNum(Trim(Request.Form("t2")),0)
	t3=FilterText(Trim(Request.Form("t3")),0)
	t4=IsNum(Trim(Request.Form("t4")),0)
	t6=FilterText(Trim(Request.Form("t5")),1)
	up1=IsNum(Trim(Request.Form("up1")),0)
	IF ID=0 Then
		sdcms.Check_lever 21
		Set Rs=Conn.Execute("select max(ordnum) from ["&Sd_Table&"] where ispic="&t2&"")
		IF Rs.Eof Then
			t5=1
		Else
			IF Rs(0)<>"" Then t5=Cint(rs(0))+1 Else t5=1 End IF
		End IF
		Rs.Close
		Set Rs=Nothing
	Else
		sdcms.Check_lever 22
	End IF
	
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select title,url,ispic,pic,ispass,ordnum,content from "&Sd_Table&""
	IF ID>0 Then 
	 sql=sql&" Where ID="&ID&""
	End IF
	rs.Open sql,conn,1,3
	IF ID=0 Then 
	  Rs.Addnew
	Else
	  Rs.Update
	End IF
	Rs(0)=Left(t0,255)
	Rs(1)=Left(t1,255)
	Rs(2)=Left(t2,255)
	Rs(3)=Left(t3,255)
	Rs(4)=Left(t4,255)
	IF ID=0 Then Rs(5)=t5
	Rs(6)=Left(t6,50)
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	IF ID=0  Then
		LogMsg="�������"
	Else
		LogMsg="�޸�����"
	End IF
	AddLog sdcms_adminname,GetIp,LogMsg&t0,0
	IF Len(up1)>0 Then
		Set sdcms_c=New sdcms_create
		sdcms_c.Create_index()
		Set sdcms_c=Nothing
	End IF
	Go("?")
End Sub

Sub Del
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	AddLog sdcms_adminname,GetIp,"ɾ�����ӣ�"&LoadRecord("title",Sd_Table,id),0
	Conn.Execute("Delete From "&Sd_Table&" where id="&id&"")
	IF Sdcms_Mode=2 Then
		Dim sdcms_c
		Set sdcms_c=New sdcms_create
		sdcms_c.Create_index()
		Set sdcms_c=Nothing
	End IF
	Go("?stype="&stype&"") 
End Sub

Sub Pass(t0)
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Conn.Execute("Update "&Sd_Table&" Set IsPass="&t0&" Where ID="&ID&"")
	AddLog sdcms_adminname,GetIp,"������ӣ�"&LoadRecord("title",Sd_Table,id),0
	IF Sdcms_Mode=2 Then
		Dim sdcms_c
		Set sdcms_c=New sdcms_create
		sdcms_c.Create_index()
		Set sdcms_c=Nothing
	End IF
	Go("?stype="&stype&"") 
End Sub

Sub Up
	Dim Rs
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim t0:t0=IsNum(Trim(Request.QueryString("t0")),0)
	Set Rs=Conn.Execute("select top 1 id,ordnum from ["&Sd_Table&"] where ordnum>"&ordnum&" and ispic="&t0&" order by ordnum  ")
	IF Not Rs.Eof Then
		Conn.Execute("Update  ["&Sd_Table&"] set ordnum="&rs(1)&" where id="&id&"")
		Conn.Execute("Update  ["&Sd_Table&"] set ordnum="&ordnum&" where id="&rs(0)&"")
	End IF
	Rs.Close
	Set Rs=Nothing
	IF Sdcms_Mode=2 Then
		Dim sdcms_c
		Set sdcms_c=New sdcms_create
		sdcms_c.Create_index()
		Set sdcms_c=Nothing
	End IF
	Go("?stype="&stype&"") 
End Sub

Sub Down
	Dim Rs
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim t0:t0=IsNum(Trim(Request.QueryString("t0")),0)
	Set Rs=Conn.Execute("select top 1 id,ordnum from ["&Sd_Table&"] where ordnum<"&ordnum&" and ispic="&t0&" order by ordnum desc")
	IF Not Rs.Eof Then
		Conn.Execute("Update  ["&Sd_Table&"] set ordnum="&rs(1)&" where id="&id&"")
		Conn.Execute("Update ["&Sd_Table&"] set ordnum="&ordnum&" where id="&rs(0)&"")
	End IF
	Rs.Close
	Set Rs=Nothing
	IF Sdcms_Mode=2 Then
		Dim sdcms_c
		Set sdcms_c=New sdcms_create
		sdcms_c.Create_index()
		Set sdcms_c=Nothing
	End IF
	Go("?stype="&stype&"") 
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��վ���Ʋ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t1.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��վ��������Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t1.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (!document.add.t3.disabled && document.add.t3.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('ͼƬ��ַ����Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t3.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>