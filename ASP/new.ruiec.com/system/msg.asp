<!--#include file="sdcms_check.asp"-->
<%
Dim Sdcms,title,Sd_Table,Sd_Table02,Sd_Table03,Stype,keyword,Publish_Where,Tj,T,Action,Classid,Page
t=IsNum(Trim(Request.QueryString("t")),0)
Action=Lcase(Trim(Request("Action")))
Classid=IsNum(Trim(Request("Classid")),0)
KeyWord=FilterText(Trim(Request("KeyWord")),0)
Page=IsNum(Trim(Request.QueryString("page")),1)
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Select Case Action
	Case "v":title="�鿴��Ϣ"
	Case Else:stype="main":title="��������"
End Select
Sd_Table="Sd_Info"
Sd_Table02="Sd_Comment"
Sd_Table03="Sd_Digg"
Sdcms_Head
IF t=0 Then
	publish_where=" Userid>=0 "
Else
	publish_where=" Userid<0 "
	title="Ͷ�����"
End IF
%>

<ul id="sdcms_sub_title">
	<li class="sub"><a<%if stype<>"main" then%> href="javascript:void(0)" onClick="selectTag('tagContent0',this)"<%end if%>><%=title%></a></li>
	<%if stype<>"main" then%>
	<li class="unsub"><a href="msg.asp">�����б�</a></li>
	<%end if%>
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "v":sdcms.Check_lever 13:view
	Case "del":sdcms.Check_lever 14:del
	Case "yd":sdcms.Check_lever 15:Changeyd
	Case "wd":sdcms.Check_lever 16:Changewd
	Case "ts":sdcms.Check_lever 17:Changets
	Case "jj":sdcms.Check_lever 18:Changejj
	Case Else:main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing
Sub Main	
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
   <form name="add" action="?t=<%=t%>&Page=<%=Page%>" method="post" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');"> 
	<tr>
	  <td width="30" class="title_bg">ѡ��</td>
	  <td width="30" class="title_bg">ID</td>
      <td class="title_bg">����</td>
      <td width="200" class="title_bg">��˾</td>
	  <td width="80" class="title_bg">�绰</td>
	  <td width="40" class="title_bg">���</td>
	  <td width="100" class="title_bg">IP</td>
	  <td width="150" class="title_bg">�����</td>
	  <td width="80" class="title_bg">����ϵͳ</td>
	  <td width="130" class="title_bg">ʱ��</td>
      <td width="100" class="title_bg">����</td>
    </tr>
	<%
	IF Classid<>0 Then tj=" And Classid In("&Get_Son_Classid(Classid)&") "
	Dim Where
	Where=""&publish_where&" And "
	
	IF Sdcms_DataType Then
		Where=Where&"(InStr(1,LCase(Title),LCase('"&keyword&"'),0)<>0 or InStr(1,LCase(id),LCase('"&keyword&"'),0)<>0) "
	Else
		Where=Where&"(title like '%"&keyword&"%' Or id like '%"&keyword&"%')"
	End IF
		
	IF Load_Cookies("sdcms_admin")=0 Then
		Dim SdcmsAdmin
		Set SdcmsAdmin=New Sdcms_Admin
		Where=Where&" And "&SdcmsAdmin.Check_Info_Lever&""
		Set SdcmsAdmin=Nothing
	End IF
	
	Where=Where&" "&tj&" "
	
	Dim P,Rs,I,Num,Url
	
	Set P=New Sdcms_Page
	With P
	.Conn=Conn
	.PageNum=Page
	.Table="sd_msg"
	.Field="*"
	.Key="ID"
	.Where=""
	.Order="addtime desc,statu,id desc"
	.PageStart=""
	End With
	On Error ReSume Next
	Set Rs=P.Show
	IF Err Then
		Num=0
		Err.Clear
	End IF
	For I=1 To P.PageSize
		IF Rs.Eof Or Rs.Bof Then Exit For
		Select Case Sdcms_Mode
			Case "2","1":Url=Rs(12)&Rs(13)
			Case Else:Url=Rs(0)
		End Select
		Dim strong,strongs
		strong = ""
		strongs = ""
		If Rs("statu")="0" Then
			strong = "<strong>"
			strongs = "</strong>"
		End If
	%>
   <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	  <td height="25" align="center"><input name="id" type="checkbox" value="<%=Rs("ID")%>"></td>
      <td align="center"><%=strong%><%=Rs("ID")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("username")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=rs("company")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("tel")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=IIF(Rs("ispass")="0","<font color='red'>δͨ��</font>","<font color='#0099CC'>��ͨ��</font>")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("addip")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("addbs")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("addos")%><%=strongs%></td>
	  <td align="center"><%=strong%><%=Rs("addtime")%><%=strongs%></td>
      <td align="center"><a href="?action=v&id=<%=rs("ID")%>">�鿴</a> <a href="?action=del&id=<%=rs("ID")%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="11" class="tdbg" align="left">
		<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label>
		<select name="action" onchange="if(this.value=='movelist'){t0.disabled=false}else{t0.disabled=true};">
			<option>������</option>
			<option value="yd">��Ϊ�Ѷ�</option>
			<option value="wd">��Ϊδ��</option>
			<option>----------</option>
			<option value="ts">ͨ�����</option>
			<option value="jj">�ܾ����</option>
			<option>----------</option>
			<option value="del">ɾ����Ϣ</option>
		</select>
		<input type="submit" class="bnt01" value="ִ��" />
	  </td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="11" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>
<%
Set P=Nothing
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Split(ID,", ")
		Dim I
		For I=0 To Ubound(ID)
			Conn.Execute("Delete From sd_msg Where ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "msg.asp"
End Sub

'------------------------
Sub Changeyd
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Split(ID,", ")
		Dim I
		For I=0 To Ubound(ID)
			Conn.Execute("UPDATE sd_msg SET statu = '1' WHERE ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "msg.asp"
End Sub

Sub Changewd
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Split(ID,", ")
		Dim I
		For I=0 To Ubound(ID)
			Conn.Execute("UPDATE sd_msg SET statu = '0' WHERE ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "msg.asp"
End Sub

Sub Changets
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Split(ID,", ")
		Dim I
		For I=0 To Ubound(ID)
			Conn.Execute("UPDATE sd_msg SET ispass = '1' WHERE ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "msg.asp"
End Sub

Sub Changejj
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		ID=Split(ID,", ")
		Dim I
		For I=0 To Ubound(ID)
			Conn.Execute("UPDATE sd_msg SET ispass = '0' WHERE ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "msg.asp"
End Sub

Sub view
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 Then
		If Request.Form("vpass")<>"" Then
			Conn.Execute("UPDATE sd_msg SET ispass = '"&Clng(Request.Form("vpass"))&"' WHERE ID = "&Clng(ID)&"")
			Go "msg.asp"
		End If

		Dim Rs
		Set Rs = Conn.Execute("SELECT * From sd_msg Where ID="&Clng(ID)&"")
		If Rs("statu")="0" Then 
			Conn.Execute("Update sd_msg Set statu='1' Where ID="&Clng(ID)&"")
		End If
%>
	<style>.input{width:300px}</style>
	<form id="myFM" method="post" action="">
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" id="tagContent1" class="table_b">
		<tr class="tdbg">
			<td align="center" width="120">ID��</td>
			<td><input value="<%=Rs("ID")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">�ա�����</td>
			<td><input value="<%=Rs("username")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">��  ˾��</td>
			<td><input value="<%=Rs("company")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">�硡����</td>
			<td><input value="<%=Rs("tel")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">��  �ݣ�</td>
			<td><textarea readonly="readonly" class="input" style="width:60%;height:150px;"><%=Rs("content")%></textarea></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">ʱ���䣺</td>
			<td><input value="<%=Rs("addtime")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">IP��</td>
			<td><input value="<%=Rs("addip")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">����ϵͳ��</td>
			<td><input value="<%=Rs("addos")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">�������</td>
			<td><input value="<%=Rs("addbs")%>" type="text" class="input" readonly="readonly" style="color:#666"></td>
		</tr>
		<tr class="tdbg">
			<td align="center" width="120">������Ϣ��</td>
			<td><textarea readonly="readonly" class="input" style="width:60%;height:150px;"><%=Rs("addother")%></textarea></td>
		</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" >
		<tr class="tdbg">
			<td width="100">&nbsp;</td>
			<td>
<%			If Rs("ispass")="0" Then %>
			<input type="hidden" name="vpass" value="1" />
			<input type="submit" class="bnt" value="ͨ�����" onclick="return confirm('ȷ��Ҫͨ����������?');" />
<%			Else %>
			<input type="hidden" name="vpass" value="0" />
			<input type="submit" class="bnt" value="�ܾ����" onclick="return confirm('ȷ��Ҫ�ܾ���������?');" />
<%			End If %>
			<a href="msg.asp">����</a>
		    </td>
		</tr>
	</table>
	</form>
<%
	Else
		Go "msg.asp"
	End IF

End Sub
%>
</div>
</body>
</html>