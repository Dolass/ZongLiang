<!--#include file="sdcms_check.asp"-->
<%
Dim Sdcms,title,Sd_Table,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Sdcms.Check_lever 2
Set Sdcms=Nothing
title="��־����"
Sd_Table="sd_log"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="?">��־����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li><li class="unsub"><a href="?action=del_all" onclick='return confirm("���Ҫɾ��?���ɻָ�!\n\n2���ڵĽ��ᱻ����!");'>�����־</a></li>
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "del":del
	Case "del_all":del_all
	Case Else:main
End Select
Db_Run
CloseDb
Sub main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <form name="add" action="?action=del" method="post"  onSubmit="return confirm('���Ҫɾ��?���ɻָ�!\n\n2���ڵĽ��ᱻ����!');">
	<tr>
	<td width="30" class="title_bg">ѡ��</td>
      <td width="60" class="title_bg">���</td>
      <td width="120" class="title_bg">�ʻ�</td>
	  <td width="*" class="title_bg">����</td>
      <td width="120" class="title_bg">IP</td>
      <td width="160" class="title_bg">����</td>
      <td width="40" class="title_bg">����</td>
    </tr>
	<%
 	Dim Page,P,Rs,i,j
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,sdcms_name,content,ip,adddate"
	.Key="ID"
	.Order="ID Desc"
	.PageStart="?page="
	End With
	On Error ReSume Next
	Set Rs=P.Show
	IF Err Then
		Err.Clear
	End IF
	For I=1 To P.PageSize
		IF Rs.Eof Or Rs.Bof Then Exit For
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	<td height="25" align="center"><input name="id"  type="checkbox" value="<%=Rs(0)%>"></td>
	 <%
	 For j=0 To 4%>
      <td <%if j<>2 then%>align="center"<%end if%>><%=Rs(j)%></td>
      <%Next%>
      <td align="center"><a href="?action=del&id=<%=Rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!\n\n�����ڵ���־���ᱻ����!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="7" class="tdbg" >
	  <input name="chkAll" type="checkbox" onclick=CheckAll(this.form) value="checkbox" id="chkall"><label for="chkall">ȫѡ</label> <input type="submit" class="bnt01" value="ɾ��">

</td>
    </tr>
	<tr>
      <td colspan="7" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	</form>
  </table>

<%
Set P=Nothing
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		Dim Sql
		Sql="Delete From "&Sd_Table&" Where Id In("&ID&") And "
		IF Sdcms_DataType Then
			Sql=Sql&"Adddate<(date()-2)"
		Else
			Sql=Sql&"Adddate<(getdate()-2)"
		End IF
		Conn.Execute(Sql)
		DbQuery=DbQuery+1
	End If
	Go "?"
End Sub

Sub Del_all
	Dim Sql
	Sql="Delete From "&Sd_Table&" Where "
	IF Sdcms_DataType Then
		Sql=Sql&"Adddate<(date()-2)"
	Else
		Sql=Sql&"Adddate<(getdate()-2)"
	End IF
	Conn.Execute(Sql)
	DbQuery=DbQuery+1
	Go "?"
End Sub
%>  
</div>
</body>
</html>