<!--#include file="sdcms_check.asp"-->
<!--#include file="../Plug/Coll_Info/Conn.asp"-->
<%
Dim sdcms,title,Sd_Table,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
sdcms.Check_lever 22
Set Sdcms=Nothing
title="�ɼ�����"
Sd_Table="Sd_Coll_Config"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="Sdcms_Coll_Config.asp">�ɼ�����</a>������<a href="Sdcms_Coll_Item.asp">�ɼ�����</a> (<a href="Sdcms_Coll_Item.asp?action=add">���</a>)������<a href="Sdcms_Coll_Filters.asp">���˹���</a> (<a href="Sdcms_Coll_Filters.asp?action=add">���</a>)������<a href="Sdcms_Coll_History.asp">��ʷ��¼</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li>	
</ul>
<div id="sdcms_right_b">
<%
Select Case Lcase(action)
Case "save":Collection_Data:Save
Case Else:Collection_Data:main
End Select
Db_Run
CloseDb


Sub Main
Dim Rs
Set Rs=Coll_Conn.Execute("select Timeout,MaxFileSize,FileExtName,UpfileDir from "&Sd_Table&" where id=1")
DbQuery=DbQuery+1
IF Rs.Eof Then
	Echo "����Ƿ��ύ����":Exit Sub
End IF
Echo Check_Add
%>
<form name="add" method="post" action="?action=save" onSubmit="return checkadd()">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr>
      <td width="120" align="center" class="tdbg">��ʱ���ã�      </td>
      <td class="tdbg"><input name="t0" type="text" class="input" value="<%=rs(0)%>" size="30">��<span>Ĭ��64�� ���128K������64�뻹���ز���(��ÿ��2K���ع���)����ʱ��</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">������С��      </td>
      <td><input name="t1" type="text" class="input" value="<%=rs(1)%>"  size="30">��<span>��λ��KB���������ļ��Ĵ�С�������������롰0��</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�������ͣ�      </td>
      <td><input name="t2" type="text" class="input" value="<%=rs(2)%>"  size="30">��<span>�ɼ������ļ�����,��ʽ:Rm|swf|rar</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">����Ŀ¼��      </td>
      <td><input name="t3" type="text" class="input" value="<%=rs(3)%>"  size="30">��<span>�ɼ�����Ŀ¼,���治��Ҫ��"/"����</span></td>
    </tr>
	<tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input name="Submit" type="submit" class="bnt" value="�� ��"></td>
    </tr>
	</table>
<%
Rs.Close
Set Rs=Nothing
End Sub

Sub Save
	Dim t0,t1,t2,t3,Rs,Sql
	t0=IsNum(Trim(Request.Form("t0")),64)
	t1=IsNum(Trim(Request.Form("t1")),0)
	t2=Trim(Request.Form("t2"))
	t3=FilterText(Trim(Request.Form("t3")),0)
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select Timeout,MaxFileSize,FileExtName,UpfileDir From "&Sd_Table&" where id=1"
	Rs.Open Sql,Coll_Conn,1,3
	Rs.update
		Rs(0)=Left(t0,50)
		Rs(1)=Left(t1,50)
		Rs(2)=Left(t2,50)
		Rs(3)=Left(t3,50)
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	Del_Cache("Get_Coll_Config")
	Alert "����ɹ���","?"
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��ʱ���ò���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t1.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('������С����Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t1.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t2.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('�������Ͳ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t2.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t3.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('����Ŀ¼����Ϊ��');"&vbcrlf
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