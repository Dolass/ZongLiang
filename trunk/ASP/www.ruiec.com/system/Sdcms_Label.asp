<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Select Case Action
	Case "add":title="�����Ƭ"
	Case "edit":title="�޸���Ƭ"
	Case Else:title="��Ƭ����"
End Select
Sd_Table="Sd_Label"
Sdcms_Head
%>

<div class="sdcms_notice"><span>���������</span><a href="?action=add">�����Ƭ</a>������<a href="?">��Ƭ����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li> 
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 25:add
	Case "edit":sdcms.Check_lever 26:add
	Case "save":save
	Case "del":sdcms.Check_lever 27:del
	Case Else:main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing
Sub Main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <form name="add" action="?action=del" method="post"  onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<tr>
	  <td width="30" class="title_bg">ѡ��</td>
      <td class="title_bg">��Ƭ����</td>
	  <td width="160" class="title_bg">˵��</td>
	  <td width="60" class="title_bg">��֤</td>
      <td width="160" class="title_bg">����</td>
      <td width="80" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,j,num
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,title,Notes,ispass,adddate"
	.Key="ID"
	.Order="ID Desc"
	.PageStart="?page="
	End With
	On Error ReSume Next
	Set Rs=P.Show
	IF Err Then
		num=0
		Err.Clear
	End IF
	For I=1 To P.PageSize
		IF Rs.Eof Or Rs.Bof Then Exit For
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	 <td height="25" align="center"><input name="id" type="checkbox" value="<%=rs(0)%>"></td>
	 <td>{sdcms_<%=rs(1)%>}</td>
	 <td><%=rs(2)%></td>
	 <td align="center"><%=IIF(rs(3)=1,"��","��")%></td>
	 <td align="center"><%=rs(4)%></td>
      <td align="center"><a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="6" class="tdbg" >
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label> <input type="submit" class="bnt01" value="ɾ��"></td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="6" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>

<%
Set P=Nothing
End Sub
Sub add
	Dim Rs,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	Check_Add
	IF ID>0 Then
		Set Rs=Conn.Execute("select title,ispass,content,notes from "&Sd_Table&" where id="&id&"")
		IF Rs.Eof then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
		End IF
	Else
		t1=1
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit='return checkadd()'>
    <tr>
      <td width="120" align="center" class="tdbg">��Ƭ���ƣ�      </td>
      <td>{sdcms_<%IF ID=0 Then%><input name="t0" type="text" class="input" id="t0" size="20"><%Else%><%=t0%><%End IF%>}<%=IIF(ID=0,"��<span>��Ƭ���ƾ���{sdcms_��ͷ�������ظ������ɸ���</span>","")%></td>
    </tr>
	<tr class="tdbg" id="c">
      <td align="center">��Ƭ���ݣ�</td>
      <td><textarea name="t2" rows="16" class="inputs" id="t2"><%=Content_Encode(t2)%></textarea></td>
   </tr>
    <tr>
      <td align="center" class="tdbg">��Ƭ˵����      </td>
      <td class="tdbg"><input name="t3" type="text" class="input" id="t3" size="20" value="<%=t3%>">
        <span>����Ϊ��̨������ʹ�ã����������壬����Ϊ��</span></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">��Ƭ���ԣ�      </td>
      <td class="tdbg"><input name="t1" id="t1" type="checkbox" value="1" <%=IIF(t1=1,"checked","")%> /><label for="t1">ͨ����֤</label></td>
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
	Dim t0,t1,t2,t3,rs,sql,LogMsg,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=Trim(Request.Form("t0"))
	t1=IsNum(Trim(Request.Form("t1")),0)
	t2=Trim(Request.Form("t2"))
	t3=FilterText(Trim(Request.Form("t3")),1)
	IF ID=0 Then sdcms.Check_lever 25 Else sdcms.Check_lever 26
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="select title,ispass,content,notes,adddate,id from "&Sd_Table&""
	IF ID=0 Then 
		Sql=Sql&" where title='"&t0&"'"
	Else
		Sql=Sql&" where id="&id&""
	End IF
	Rs.Open Sql,Conn,1,3
	IF ID=0 Then 
		IF Not Rs.Eof Then
			Echo "����Ƭ���Ѵ��ڣ��뻻������":Died
		End IF
	End IF
	IF ID=0 Then 
		Rs.Addnew
	Else
		Rs.Update
	End IF
	IF ID=0 Then 
	Rs(0)=left(t0,50)
	End IF
	Rs(1)=t1
	Rs(2)=t2
	Rs(3)=left(t3,50)
	IF ID=0 Then Rs(4)=Dateadd("h",Sdcms_TimeZone,Now())
	Rs.Update
	IF ID=0 Then LogMsg="�����Ƭ" Else LogMsg="�޸���Ƭ"
	AddLog sdcms_adminname,GetIp,LogMsg&"{sdcms_"&rs(0)&"}",0
	Rs.Close
	Set Rs=Nothing
	Del_Cache("Load_Freelabel")
	Go("?")
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Re(Id," ","") 
		AddLog sdcms_adminname,GetIp,"ɾ����Ƭ�����Ϊ"&ID,0
		Conn.Execute("Delete From "&Sd_Table&" Where Id In("&id&")")
	End IF
	Del_Cache("Load_Freelabel")
	Go("?")
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	IF Action="add" Then
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��Ƭ���Ʋ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	End IF
	Check_Add=Check_Add&"	if (document.add.t2.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��Ƭ���ݲ���Ϊ��');"&vbcrlf
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