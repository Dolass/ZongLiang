<!--#include file="sdcms_check.asp"-->
<!--#include file="../Plug/Coll_Info/Conn.asp"-->
<%
Dim sdcms,title,Sd_Table,stype,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Sdcms.Check_lever 23
Set Sdcms=Nothing
title="��ʷ��¼"
Sd_Table="Sd_Coll_History"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="Sdcms_Coll_Config.asp">�ɼ�����</a>������<a href="Sdcms_Coll_Item.asp">�ɼ�����</a> (<a href="Sdcms_Coll_Item.asp?action=add">���</a>)������<a href="Sdcms_Coll_Filters.asp">���˹���</a> (<a href="Sdcms_Coll_Filters.asp?action=add">���</a>)������<a href="Sdcms_Coll_History.asp">��ʷ��¼</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li><li class="unsub"><a href="?action=del_all" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>��ռ�¼</a></li>	
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "del":Collection_Data:del
	Case "del_all":Collection_Data:del_all
	Case Else:Collection_Data:main
End Select
Db_Run
CloseDb

Sub main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b" id="tagContent0">
    <form name="add" action="?action=del" method="post" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<tr>
      <td width="30" class="title_bg">ѡ��</td>
      <td width="100" class="title_bg">��Ŀ����</td>
	  <td class="title_bg">����</td>
	  <td width="140" class="title_bg">����</td>
	  <td width="100" class="title_bg">��Ŀ</td>
      <td width="40" class="title_bg">��Դ</td>
	  <td width="40" class="title_bg">���</td>
      <td width="40" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,num,rs1
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Coll_Conn
	.PageNum=Page
	.PageSize=20
	.Table=Sd_Table
	.Field="id,ItemID,Title,Adddate,ClassID,NewsUrl,Result"
	.Key="ID"
	.Where=""
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
	  <td align="center"><%IF Rs(1)=0 Then%>δָ��<%Else%><%Set Rs1=Coll_Conn.Execute("Select ItemName From Sd_Coll_Item Where Id="&Clng(Rs(1))&""):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "��������":End IF%><%End IF%></td>
	  <td><%=Rs(2)%></td>
	  <td align="center"><%=Rs(3)%></td>
	  <td align="center"><%IF Rs(4)=0 Then%>δָ��<%Else%><%Set Rs1=Conn.Execute("Select Title From Sd_Class Where Id="&Clng(Rs(4))&""):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "��������":End IF%><%End IF%></td>
	  <td align="center"><a href="<%=Rs(5)%>" target="_blank">�鿴</a></td>
	  <td align="center"><%Select Case Rs(6):Case 1:Echo "��":Case Else:Echo "��":End Select%></td>
      <td align="center"><a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="8" class="tdbg" >
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label>              
      <input name="submit" type="submit" class="bnt01" value="ɾ��">

</td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="8" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>
  
<%
Set P=Nothing
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,"ɾ���ɼ���ʷ��¼�����Ϊ"&id,0
		Coll_conn.Execute("Delete From "&Sd_Table&" where id in("&id&")")
	End If
	Go "?"
End Sub

Sub Del_all
	AddLog sdcms_adminname,GetIp,"ɾ��ȫ���ɼ���ʷ��",0
	Coll_Conn.Execute("Delete From "&Sd_Table&"")
	Go "?"
End Sub
%>  
</div>
</body>
</html>