<!--#include file="sdcms_check.asp"-->
<!--#include file="../Plug/Coll_Info/Conn.asp"-->
<%
Dim sdcms,title,Sd_Table,Action
Action=Lcase(Trim(Request("Action")))
Set sdcms=new Sdcms_Admin
sdcms.Check_admin
Select Case action
	Case "add":title="��ӹ���"
	Case "edit":title="�޸Ĺ���"
	Case Else:title="���˹���"
End Select
Sd_Table="Sd_Coll_Filters"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="Sdcms_Coll_Config.asp">�ɼ�����</a>������<a href="Sdcms_Coll_Item.asp">�ɼ�����</a> (<a href="Sdcms_Coll_Item.asp?action=add">���</a>)������<a href="Sdcms_Coll_Filters.asp">���˹���</a> (<a href="Sdcms_Coll_Filters.asp?action=add">���</a>)������<a href="Sdcms_Coll_History.asp">��ʷ��¼</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li>
	
</ul>
<div id="sdcms_right_b">
<%
Collection_Data
Select Case Lcase(action)
	Case "add":sdcms.Check_lever 21:add
	Case "edit":sdcms.Check_lever 22:add
	Case "save":save
	Case "del":sdcms.Check_lever 23:del
	Case "pass":sdcms.Check_lever 22:pass(1)
	Case "nopass":sdcms.Check_lever 22:pass(0)
	Case Else:main
End Select
Db_Run
closedb
Set sdcms=Nothing
Sub main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b" id="tagContent0">
    <form name="add" action="?" method="post"  onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<tr>
      <td width="30" class="title_bg">ѡ��</td>
      <td class="title_bg">��������</td>
	  <td width="100" class="title_bg">����</td>
	  <td width="80" class="title_bg">����</td>
	  <td width="100" class="title_bg">������Ŀ</td>
      <td width="40" class="title_bg">״̬</td>
      <td width="120" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,num,rs1
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Coll_Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,FilterName,FilterObject,FilterType,ItemID,Flag"
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
	 <td height="25" align="center"><input name="id"  type="checkbox" value="<%=rs(0)%>"></td>
      <td><%=Rs(1)%></td>
	  <td align="center"><%Select Case Rs(2):Case "1":Echo "�������":Case "2":Echo "���Ĺ���":End Select%></td>
	  <td align="center"><%Select Case Rs(3):Case "1":Echo "���滻":Case "2":Echo "�߼�����":End Select%></td>
	  <td align="center"><%IF Rs(4)=0 Then%>δָ��<%Else%><%Set Rs1=Coll_Conn.Execute("Select ItemName From Sd_Coll_Item Where Id="&Clng(Rs(4))&""):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "��������":End IF%><%End IF%></td>
	  <td align="center"><%=IIF(Rs(5),"��","<b>��</b>")%></td>
      <td align="center"><%IF Rs(5)=1 Then%><a href="?action=noPass&id=<%=rs(0)%>&t=0">����</a><%Else%><a href="?action=Pass&id=<%=rs(0)%>&t=1">����</a><%End IF%> <a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="7" class="tdbg" >
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label>  
              <select name="action">
			  <option>������</option>
			  <option value="Pass">����</option>
			  <option value="NoPass">����</option>
			  <option value="Del">ɾ��</option>
			  </select> 
             
      <input name="submit" type="submit" class="bnt01" value="ִ��">

</td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="7" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>
  
<%
Set P=Nothing
End Sub

Sub Add
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID>0 Then
		Dim Rs,Rs1,I
		Set Rs=coll_conn.execute("select FilterName,ItemID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag from "&Sd_Table&" where id="&id&"")
		DbQuery=DbQuery+1
		IF Rs.Eof Then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3,t4,t5,t6,t7,t8
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
			t4=Rs(4)
			t5=Rs(5)
			t6=Rs(6)
			t7=Rs(7)
			t8=Rs(8)
		End IF
	Else
		t3=1
		t8=t3
	End IF
	check_info
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
    <tr>
      <td width="120" align="center" class="tdbg">�������ƣ�</td>
      <td class="tdbg"><input name="t0" type="text" class="input" id="t0" size="40" value="<%=t0%>"></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">������Ŀ��</td>
      <td class="tdbg"><select name="t1"><option value="">��ѡ����Ŀ</option><%Set Rs1=Coll_Conn.Execute("Select id,ItemName From Sd_Coll_Item Order By Id Desc"):DbQuery=DbQuery+1:While Not Rs1.Eof%><option value="<%=Rs1(0)%>" <%=IIF(t1=Rs1(0),"selected","")%>><%=Rs1(1)%></option><%Rs1.MoveNext:Wend%></select></td>
    </tr>
    <tr>
      <td align="center" class="tdbg">���˶���</td>
      <td class="tdbg"><select name="t2"><option value="1" <%=IIF(t2=1,"selected","")%>>�������</option><option value="2" <%=IIF(t2=2,"selected","")%>>���Ĺ���</option></select></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">�������ͣ�</td>
      <td class="tdbg"><input name="t3" type="radio" value="1" <%=IIF(t3=1,"checked","")%> onclick="<%For I=2 To 3%>$('#f0<%=I%>')[0].style.display='none';<%Next%>$('#f01')[0].style.display='';this.form.t4.disabled=false;" id="t3_0"><label for="t3_0">���滻</label> <input name="t3" type="radio" value="2" onclick="<%For I=2 To 3%>$('#f0<%=I%>')[0].style.display='';<%Next%>$('#f01')[0].style.display='None';this.form.t4.disabled=true;" <%=IIF(t3=2,"checked","")%> id="t3_1"><label for="t3_1">�߼�����</label></td>
    </tr>
	<tr id="f01" <%=IIF(t3=2,"class=""dis""","")%>>
      <td width="120" align="center" class="tdbg">�� �� �ݣ�</td>
      <td class="tdbg"><textarea name="t4" rows="5" class="inputs" id="t4"><%=Content_Encode(t4)%></textarea></td>
    </tr>
	<tr id="f02" <%=IIF(t3=1,"class=""dis""","")%>>
      <td width="120" align="center" class="tdbg">��ʼ��־��</td>
      <td class="tdbg"><textarea name="t5" rows="5" class="inputs"><%=Content_Encode(t5)%></textarea></td>
    </tr>
	<tr id="f03" <%=IIF(t3=1,"class=""dis""","")%>>
      <td width="120" align="center" class="tdbg">������־��</td>
      <td class="tdbg"><textarea name="t6" rows="5" class="inputs"><%=Content_Encode(t6)%></textarea></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">�� �� Ϊ��</td>
      <td class="tdbg"><textarea name="t7" rows="5" class="inputs"><%=Content_Encode(t7)%></textarea></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">״����̬��</td>
      <td class="tdbg"><input name="t8" type="checkbox" value="1" <%=IIF(t8=1,"checked","")%> id="t8_0"><label for="t8_0">����</label></td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="�� ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Save
	Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,Rs,Sql,LogMsg
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=Trim(Request.Form("t0"))
	t1=Trim(Request.Form("t1"))
	t2=Trim(Request.Form("t2"))
	t3=Trim(Request.Form("t3"))
	t4=Trim(Request.Form("t4"))
	t5=Trim(Request.Form("t5"))
	t6=Trim(Request.Form("t6"))
	t7=Trim(Request.Form("t7"))
	t8=Trim(Request.Form("t8"))
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="select FilterName,ItemID,FilterObject,FilterType,FilterContent,FisString,FioString,FilterRep,Flag,id From "&Sd_Table
	DbQuery=DbQuery+1
	IF ID>0 then 
		Sql=Sql&" where id="&id&""
	End if
	Rs.Open Sql,Coll_Conn,1,3
	
	IF ID=0 Then 
	  rs.Addnew
	Else
	  rs.Update
	End IF
	rs(0)=Left(t0,50)
	rs(1)=IsNum(t1,0)
	rs(2)=IsNum(t2,1)
	rs(3)=IsNum(t3,0)
	rs(4)=t4
	rs(5)=t5
	rs(6)=t6
	rs(7)=t7
	rs(8)=IsNum(t8,0)
	rs.Update
	IF ID=0 Then
		LogMsg="��Ӳɼ�����"
	Else
		LogMsg="�޸Ĳɼ�����"
	End IF
	Del_Cache("Get_Coll_Filters")
	AddLog sdcms_adminname,GetIp,LogMsg&ID,0
	Rs.Close
	Set Rs=Nothing
	Echo "����ɹ�"
End Sub

Sub Del 
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		Del_Cache("Get_Coll_Filters")
		ID=Re(ID," ","")
		AddLog sdcms_adminname,GetIp,"ɾ���ɼ�������Ŀ�����Ϊ"&id,0
		Coll_conn.Execute("Delete from "&Sd_Table&" where id in("&id&")")
	End IF
	Go("?") 
End Sub

Sub Pass(t0)
	Dim ID:ID=Trim(Request("ID"))
	ID=Replace(ID," ","")
	IF Len(ID)>0 Then
		Coll_Conn.Execute("Update "&Sd_Table&" Set flag="&t0&" where id in("&id&")")
	End IF
	Go("?") 
End Sub

Sub check_info
	Echo"<script>"
	Echo"	function checkadd()"
	Echo"	{"
	Echo"	if (document.add.t0.value=='')"
	Echo"	{"
	Echo"	alert('�������Ʋ���Ϊ��');"
	Echo"	document.add.t0.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	if (document.add.t1.value=='')"
	Echo"	{"
	Echo"	alert('��ѡ����Ŀ');"
	Echo"	document.add.t1.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	if (!document.add.t4.disabled && document.add.t4.value=='')"
	Echo"	{"
	Echo"	alert('Ҫ�滻�����ݲ���Ϊ��');"
	Echo"	document.add.t4.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	if (document.add.t4.disabled && document.add.elements.t5.value =='')"
	Echo"	{"
	Echo"	alert('��ʼ��־����Ϊ��');"
	Echo"	document.add.t5.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	if (document.add.t4.disabled && document.add.elements.t6.value =='')"
	Echo"	{"
	Echo"	alert('������־����Ϊ��');"
	Echo"	document.add.t6.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	if (document.add.t7.value=='')"
	Echo"	{"
	Echo"	alert('�滻�Ľ������Ϊ��');"
	Echo"	document.add.t7.focus();"
	Echo"	return false"
	Echo"	}"
	Echo"	}"
	Echo"</script>"
End Sub
%>  
</div>
</body>
</html>