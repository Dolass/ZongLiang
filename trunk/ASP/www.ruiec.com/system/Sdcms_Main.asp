<!--#include file="sdcms_check.asp"-->
<%
Dim Sdcms,Action:Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Set Sdcms=Nothing
Sdcms_Head
Main
Db_Run
Sub Main
%>
 
<br>
<ul id="sdcms_sub_title">
	<li class="sub">��������Ϣ</li>
</ul>
<div id="sdcms_right_b">
<table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <tr>
      <td height="25" align="center" class="tdbg01">Adodb.Stream</td>
      <td align="center" class="tdbg01">Scripting.FileSystemObject</td>
      <td align="center" class="tdbg01">Microsoft.XMLHTTP</td>
	  <td align="center" class="tdbg01">Persits.Jpeg</td>
      <td align="center" class="tdbg01">JMail.SmtpMail</td>
	  <td align="center" class="tdbg01">�ű���ʱ</td>
    </tr>
	<tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="20" align="center"><%=IIF(IsObjInstalled("Adodb.Stream"),"��","��")%></td>
      <td align="center"><%=IIF(IsObjInstalled("Scripting.FileSystemObject"),"��","��")%></td>
	  <td align="center"><%=IIF(IsObjInstalled("Microsoft.XMLHTTP"),"��","��")%></td>
	  <td align="center"><%=IIF(IsObjInstalled("Persits.Jpeg"),"��","��")%></td>
      <td align="center"><%=IIF(IsObjInstalled("JMail.SMTPMail"),"��","��")%></td>
	  <td align="center"><%=Server.ScriptTimeout%> ��</td>
    </tr>
  </table>
</div>


<br>
<ul id="sdcms_sub_title">
	<li class="sub">��ҵ��Ϣ</li>
</ul>
<div id="sdcms_right_b">
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td width="100" height="22" align="center">��˾���ƣ�</td>
      <td>����Դ����Ƽ����޹�˾</td>
    </tr> 
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">��ǰ��վ��</td>
      <td><a href="http://www.sdcms.cn">www.ruiec.com</a></td>
    </tr>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">��ǰ�汾��</td>
      <td>1.31</td>
    </tr>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">���վ�㣺</td>
      <td><a href="http://www.ruiec.cn" target="_blank">��������</a>&nbsp;&nbsp;<a href="http://www.ruiec.net" target="_blank">����IDC</a></td>
    </tr>
  </table>
</div>

</body>
</html>
<%End Sub%>