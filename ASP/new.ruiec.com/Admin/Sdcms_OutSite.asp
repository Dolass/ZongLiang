<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Select Case Action
	Case "add":title="��ӵ���"
	Case "edit":title="�޸ĵ���"
	Case Else:title="���ù���"
End Select
Sd_Table="Sd_OutSite"
Sdcms_Head
%>

<div class="sdcms_notice"><span>���������</span><a href="?action=add">��ӵ���</a>������<a href="?">���ù���</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li> 
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 18:add
	Case "edit":add
	Case "save":save
	Case "del":sdcms.Check_lever 20:del
	Case "getcode":GetCode
	Case "view":view
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
      <td class="title_bg">����˵��</td>
	  <td width="60" class="title_bg">����ʱ��</td>
	  <td width="60" class="title_bg">��֤</td>
      <td width="140" class="title_bg">����</td>
      <td width="140" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,j,num
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,title,CacheTime,IsPass,adddate"
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
	 <td><%=rs(1)%></td>
	 <td align="center" class="tdbg"><%=rs(2)%> <span>��</span></td>
	 <td align="center"><%=IIF(rs(3)=1,"��","��")%></td>
	 <td align="center"><%=rs(4)%></td>
      <td align="center"><a href="?action=getcode&id=<%=rs(0)%>">����</a> <a href="?action=view&id=<%=rs(0)%>">Ԥ��</a> <a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
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

Sub Add
	Dim Rs,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	Check_Add
	IF ID>0 Then
		Set Rs=Conn.Execute("Select title,CacheTime,Loop_Content,ispass From "&Sd_Table&" Where Id="&id&"")
		IF Rs.Eof then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3,t4,t5
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
		End IF
	Else
		t1=60
		t3=1
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
    <tr class="tdbg">
      <td width="120" align="center">����˵����      </td>
      <td><input name="t0" type="text" class="input" id="t0" size="20" value="<%=t0%>" />��<span>����˵��������Ϊ��̨������ʹ��</span></td>
    </tr>
	<tr class="tdbg">
	    <td align="center">����ʱ�䣺</td>
	    <td><input name="t1" type="text" class="input" id="t1" size="20" value="<%=t1%>" onKeypress="if(event.keyCode<45||event.keyCode>57)event.returnValue=false;" />��<span>���ݻ���ʱ�䣬��λΪ��</span></td>
	</tr>
	<tr class="tdbg">
      <td align="center">��ǩ���ݣ�</td>
      <td><textarea name="t2" rows="16" class="inputs" id="t3"><%=Content_Encode(t2)%></textarea><span>ֱ��ʹ��ϵͳ��ǩ����</span></td>
   </tr>
   <tr class="tdbg">
        <td align="center">�������ԣ�      </td>
        <td><input name="t3" id="t3" type="checkbox" value="1" <%=IIF(t3=1,"checked","")%> /><label for="t3">ͨ����֤</label></td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��������"> <input type="button" onClick="history.go(-1)" class="bnt" value="��������"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub GetCode
	Dim Rs,ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Set Rs=Conn.Execute("Select Id from "&Sd_Table&" Where Id="&id&"")
	DbQuery=DbQuery+1
	IF Rs.Eof Then
		Echo "����Ƿ��ύ����":Exit Sub
	End IF
	Rs.Close
	Set Rs=Nothing
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr class="tdbg">
    <tr class="tdbg">
      <td align="center">���룺</td>
      <td><textarea name="get_c" rows="2" class="inputs"><script src="<%=Sdcms_WebUrl%><%=sdcms_root%>Plug/OutSite.Asp?ID=<%=ID%>" language="javascript"></script></textarea></td>
    </tr>
    <tr class="tdbg">
      <td>&nbsp;</td>
      <td><input  type="button"   class="bnt" value="����" onClick="CopyUrl(get_c);"> <input name="Submit22" type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
</table>
<%
End Sub

Sub View
	Dim Rs,ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Set Rs=Conn.Execute("Select Id from "&Sd_Table&" Where Id="&id&"")
	DbQuery=DbQuery+1
	IF Rs.Eof Then
		Echo "����Ƿ��ύ����":Exit Sub
	End IF
	Rs.Close
	Set Rs=Nothing
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr class="tdbg">
      <td width="120" align="center">���ô��룺</td>
      <td><textarea name="get_c" rows="2" class="inputs"><script src="<%=Sdcms_WebUrl%>/Plug/OutSite.Asp?ID=<%=ID%>" language="javascript"></script></textarea></td>
    </tr>
    <tr class="tdbg">
      <td width="120" align="center">Ч��Ԥ����</td>
      <td><script src="<%=Sdcms_WebUrl%>/Plug/OutSite.Asp?ID=<%=ID%>" language="javascript"></script></td>
    </tr>
    <tr class="tdbg">
      <td>&nbsp;</td>
      <td><input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
</table>
<%
End Sub

Sub Save
	Dim t0,t1,t2,t3
	Dim Rs,sql,ID
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=Trim(Request.Form("t0"))
	t1=IsNum(Trim(Request.Form("t1")),0)
	t2=Request.Form("t2")
	t3=IsNum(Trim(Request.Form("t3")),0)
	IF ID=0 Then sdcms.Check_lever 18 Else sdcms.Check_lever 19
	
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select title,CacheTime,Loop_Content,ispass From "&Sd_Table&""
	IF ID<>0 Then 
		Sql=Sql&" Where Id="&Id&""
	End IF
	Rs.Open Sql,Conn,1,3

	IF ID=0 Then 
		Rs.Addnew
	Else
		Rs.Update
	End IF
	Rs(0)=left(t0,50)
	Rs(1)=t1
	Rs(2)=t2
	Rs(3)=t3
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Go("?")
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Re(Id," ","") 
		AddLog sdcms_adminname,GetIp,"ɾ�����ã����Ϊ"&ID,0
		Conn.Execute("Delete From "&Sd_Table&" Where Id In("&id&")")
	End IF
	Go("?")
End Sub

Function Check_Add
	Check_Add="<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('����˵������Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t1.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('����ʱ�䲻��Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t1.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t2.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��ǩ���ݲ���Ϊ��');"&vbcrlf
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