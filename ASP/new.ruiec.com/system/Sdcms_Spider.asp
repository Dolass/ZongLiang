<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
sdcms.Check_lever 23
Set Sdcms=Nothing
Title="֩������"
Sd_Table="Sd_Spider"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="?">֩������</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li><li class="unsub"><a href="?action=del_all" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>��ռ�¼</a></li>
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

Sub Main
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
  <tr onmouseover="this.bgColor='#EEFEED';" onmouseout="this.bgColor='#ffffff';" bgcolor='#ffffff'>
    <td height="20" align="center" class="tdbg01">֩��</td>
    <td align="center" bgcolor="#EEFEED" class="tdbg01">�������ʱ��</td>
    <td align="center" class="tdbg01">�����ô���</td>
	<td align="center" class="tdbg01">����</td>
  </tr>
  <%Dim Rs:Set Rs=Conn.Execute("Select title,Lastdate,hits From "&Sd_Table&""):DbQuery=DbQuery+1:IF Rs.Eof Then%>
  <tr bgcolor='#ffffff'>
    <td height="25" colspan="4" align="center">�������ü�¼</td>
  </tr>
  <%End IF:While Not Rs.Eof%>
  <tr onmouseover="this.bgColor='#EEFEED';" onmouseout="this.bgColor='#ffffff';" bgcolor='#ffffff'>
    <td height="25" align="center"><%=Rs(0)%></td>
    <td align="center"><%=Rs(1)%></td>
    <td align="center"><%=Rs(2)%></td>
	<td align="center"><a href="?action=del&id=<%=Rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!\n2���ڵĽ��ᱻ����!");'>ɾ��</a></td>
  </tr>
  <%Rs.MoveNext:Wend:Rs.Close:Set Rs=Nothing%>
</table>
<%
End Sub

Sub Del
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Conn.Execute("Delete From "&Sd_Table&" where id in("&id&")")
	Go("?")
End Sub

Sub Del_All
	Conn.Execute("Delete From "&Sd_Table&"")
	Go("?")
End Sub
%>
</div>
</body>
</html>