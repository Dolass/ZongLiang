<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action,stype
Action=Lcase(Trim(Request("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Select Case Action
	Case "add":title="�������"
	Case "edit":title="�޸�����"
	Case Else:stype="main":title="��������"
End Select
Sd_Table="sd_search"
Sdcms_Head
%>

<div class="sdcms_notice"><span>���������</span><a href="?action=add">�������</a>������<a href="?">��������</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li><%IF Len(stype)>0 Then%><li class="unsub"><a href="?action=del_all" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>��ռ�¼</a></li><%End IF%>
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 18:add
	Case "edit":sdcms.Check_lever 19:add
	Case "pass":sdcms.Check_lever 19:pass(1)
	Case "nopass":sdcms.Check_lever 19:pass(0)
	Case "ontop":sdcms.Check_lever 19:ontop(1)
	Case "notop":sdcms.Check_lever 19:ontop(0)
	Case "save":save
	Case "del":sdcms.Check_lever 20:del
	Case "del_all":sdcms.Check_lever 20:del_all
	Case else:main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing
Sub Main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <form name="add" action="?" method="post"  onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<tr>
	  <td width="30" class="title_bg">ѡ��</td>
      <td class="title_bg">�ؼ���</td>
      <td width="80" class="title_bg">����</td>
      <td width="120" class="title_bg">��֤/�ö�</td>
      <td width="120" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,num
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,title,hits,ispass,ontop"
	.Key="ID"
	.Order="ontop desc,id desc"
	.PageStart="?page="
	End With
	On Error ReSume Next
	Set Rs=P.Show
	IF Err Then
		Num=0
		Err.Clear
	End IF
	For I=1 To P.PageSize
		IF Rs.Eof Or Rs.Bof Then Exit For 
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	 <td height="25" align="center"><input name="id" type="checkbox" value="<%=rs(0)%>"></td>
	  <td><a href="../Search/?/<%=rs(1)%>" title="<%=rs(1)%>"><%=rs(1)%></a></td>
	  <td align="center"><%=rs(2)%></td>
	  <td align="center"><%=IIF(Rs(3)=1,"<b>��</b>","��")%>/<%=IIF(Rs(4)=1,"<b>��</b>","��")%></td>
      <td align="center"><a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="5" class="tdbg" >
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label>  
              <select name="action">
			  <option>������</option>
			  <option value="ontop">��Ϊ�ö�</option>
			  <option value="pass">ͨ����֤</option>
              <optgroup></optgroup>
			  <option value="notop">ȡ���ö�</option>
			  <option value="nopass">ȡ����֤</option>
              <optgroup></optgroup>
			  <option value="del">ɾ����¼</option>
			  </select> 
             
      <input type="submit" class="bnt01" value="ִ��">

</td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="5" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>

<%
Set P=Nothing
End Sub

Sub Add
	Dim Rs
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID>0 Then
		Set Rs=Conn.Execute("select title,hits,ispass,ontop from "&Sd_Table&" where id="&id&"")
		DbQuery=DbQuery+1
		IF Rs.Eof Then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
		End IF
		Rs.Close
		Set Rs=Nothing
	Else
		t2=1:t3=0
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
   <tr>
      <td width="120" align="center" class="tdbg">�ؼ��֣�      </td>
      <td class="tdbg"><input name="t0" type="text" class="input" id="t0" size="40" value="<%=t0%>"></td>
    </tr>
    <tr class="tdbg">
      <td align="center">�ˡ�����</td>
      <td><input name="t1" type="text" class="input" value="<%=t1%>" id="t1" size="40" onKeyUp="value=value.replace(/[^\d]/g,'');"  onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''));"></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�����ԣ�</td>
      <td><input name="t2" type="checkbox" value="1" <%=IIF(t2=1,"checked","")%> id="t2_0"><label for="t2_0">ͨ��</label> <input name="t3" type="checkbox" value="1" <%=IIF(t3=1,"checked","")%> id="t3_0"><label for="t3_0">�ö�</label></td>
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
	Dim t0,t1,t2,t3,rs,sql,LogMsg
	t0=FilterText(Trim(Request.Form("t0")),1)
	t1=IsNum(Trim(Request.Form("t1")),0)
	t2=IsNum(Trim(Request.Form("t2")),0)
	t3=IsNum(Trim(Request.Form("t3")),0)
	IF ID=0 Then sdcms.Check_lever 18 Else sdcms.Check_lever 19
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="select title,hits,ispass,ontop,adddate From "&Sd_Table&" Where "
	IF ID=0 then
		Sql=Sql&" title='"&t0&"'"
	Else
		Sql=Sql&" id="&ID
	End IF
	Rs.Open Sql,Conn,1,3
	IF ID=0 then
		IF Not Rs.Eof Then
			Echo "�ùؼ��������Ѵ��ڣ��뻻�����ԣ�":Exit Sub
		End IF
		Rs.Addnew
	Else
		Rs.Update
	End IF
		rs(0)=Left(t0,50)
		rs(1)=t1
		rs(2)=t2
		rs(3)=t3
		IF ID=0 Then Rs(4)=Dateadd("h",Sdcms_TimeZone,Now())
	Rs.Update
	Rs.Close
	IF ID=0 Then LogMsg="��������ؼ���" Else LogMsg="�޸������ؼ���"
	AddLog sdcms_adminname,GetIp,LogMsg&t0,0
	Go("?") 
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,"ɾ�������ؼ��֣����Ϊ"&id,0
		ID=Re(ID," ","")
		Conn.Execute("Delete From "&Sd_Table&" where id in("&ID&")")
	End if
	Go("?")
End Sub

Sub Del_All
	AddLog sdcms_adminname,GetIp,"ɾ��ȫ�������ؼ���",0
	Conn.Execute("Delete From "&Sd_Table&"")
	Go("?")
End Sub

Sub Pass(t0)
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		ID=Re(ID," ","")
		Conn.Execute("Update "&Sd_Table&" Set IsPass="&t0&" where id in("&id&")")
	End if
	Go("?")
End Sub

Sub Ontop(t0)
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(id)>0 Then
		ID=Re(ID," ","")
		Conn.Execute("Update "&Sd_Table&" Set Ontop="&t0&" where id in("&id&")")
	End if
	Go("?")
End Sub
 
Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('�ؼ������Ʋ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>