<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action
Action=Lcase(Trim(Request("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
title="Tags����"
Sd_Table="sd_tags"
Sdcms_Head
%>

<div class="sdcms_notice"><span>���������</span><a href="?">Tags����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li> 
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "del":sdcms.Check_lever 20:del
	Case "edit":sdcms.Check_lever 19:edit
	Case "save":sdcms.Check_lever 19:save
	Case Else:main
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
      <td width="*" class="title_bg">����</td>
      <td width="100" class="title_bg">ʹ�ô���</td>
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
	.Field="id,title,followid,tag_font,tag_size,tag_color"
	.Key="ID"
	.Order="ID Desc"
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
	 <td height="25" align="center"><input name="id"  type="checkbox" value="<%=Rs(0)%>"></td>
	  <td><span style="font:<%=Rs(4)%> <%=Rs(3)%>;color:<%=Rs(5)%>"><%=Rs(1)%></span></td>
	  <td align="center"><%=Rs(2)%></td>
	  <td align="center"><a href="?action=edit&id=<%=Rs(0)%>">�༭</a> <a href="?action=del&id=<%=Rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="4" class="tdbg" >
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" ><label for="chkall">ȫѡ</label>  
              <select name="action">
			  <option value="del">ɾ��</option>
			  </select> 
             
      <input type="submit" class="bnt01" value="ִ��"></td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="4" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>

<%
Set P=Nothing
End Sub

Sub Edit
Dim Rs,i,all_color,k
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
Set Rs=Conn.Execute("select id,title,tag_font,tag_size,tag_color from "&Sd_Table&" where id="&id&"")
DbQuery=DbQuery+1
IF Rs.Eof Then
	Echo "����Ƿ��ύ����":Exit Sub
End IF
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit='return checkadd()'>
    <tr>
      <td width="120" align="center" class="tdbg">TagsЧ����      </td>
      <td class="tdbg"><span id="tags" style="font:<%=rs(3)%> <%=rs(2)%>;color:<%=rs(4)%>"><%=rs(1)%></span></td>
    </tr>
    <tr class="tdbg">
      <td align="center">����ѡ��      </td>
      <td><select name="t0" id="t0" onchange="$('#tags')[0].style.fontFamily=this.value;">
        <option value="����" <%=IIF(rs(2)="����","selecte","")%>>����</option>
        <option value="����" <%=IIF(rs(2)="����","selecte","")%>>����</option>
        <option value="�ź�" <%=IIF(rs(2)="�ź�","selecte","")%>>�ź�</option>
      </select>
      </td>
    </tr>
    <tr class="tdbg">
      <td align="center">�����С��      </td>
      <td><select name="t1" id="t1" onchange="$('#tags')[0].style.fontSize=this.value;">
	  <%For i=12 to 20 step 2%>
          <option value="<%=i%>px" <%=IIF(rs(3)=i&"px","selected","")%>><%=i%>px</option>
	  <%next%>
      </select>
      </td>
    </tr>
	<tr class="tdbg">
      <td align="center">������ɫ��      </td>
      <td><select name="t2" id="t2" onchange="$('#tags')[0].style.color=this.value;">
	  <%all_color="#000|#ff0|#f00|#fff|#00f|#0f0"
	  all_color=split(all_color,"|")
	  for k=0 to ubound(all_color)%>
          <option style="background:<%=all_color(k)%>;color:#333;" value="<%=all_color(k)%>" <%=IIF(rs(4)=all_color(k),"selected","")%>><%=all_color(k)%></option>
	  <%next%>
      </select>
      </td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��������"> <input type="button" onClick="history.go(-1)" class="bnt" value="��������"></td>
    </tr>
	</form>
  </table>
<%
Rs.Close
Set Rs=Nothing
End Sub

Sub Save
	Dim t0,t1,t2,Rs,Sql
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=Trim(Request.Form("t0"))
	t1=Trim(Request.Form("t1"))
	t2=Trim(Request.Form("t2"))
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select tag_font,tag_size,tag_color From "&Sd_Table&" Where id="&id&""
	rs.Open Sql,conn,1,3
	rs.Update
	rs(0)=Left(t0,4)
	rs(1)=Left(t1,4)
	rs(2)=Left(t2,4)
	rs.Update
	rs.Close
	Set Rs=Nothing
	Go("?")
End Sub
 
Sub Del
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		ID=Re(ID," ",""):AddLog sdcms_adminname,GetIp,"ɾ��Tags�����Ϊ"&id,0
		Conn.Execute("Delete From "&Sd_Table&" Where Id In("&ID&")")
	End if
	Go("?")
End Sub
%>  
</div>
</body>
</html>