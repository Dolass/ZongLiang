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
<div class="sdcms_notice" id="notice"><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="80"><span>�ٷ�֪ͨ��</span></td>
    <td width="*" align="left" id="sdcms_notice">Loading</td>
	<td height="25" align="right"><img src="images/close.gif" onclick="$('#notice').toggle('slow')" class="hand" hspace="6" vspace="6"></td>
  </tr>
</table>
</div>
 
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
	<li class="sub">�ٷ��Ƽ�</li>
</ul>
<div id="sdcms_right_b">
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr bgcolor='#ffffff'>
      <td height="108" valign="top" id="sdcms_nice">Loading</td>
      </tr>
  </table>
</div>

<br>
<ul id="sdcms_sub_title">
	<li class="sub">��Ȩ��Ϣ</li>
</ul>
<div id="sdcms_right_b">
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td width="100" height="22" align="center">�������ƣ�</td>
      <td>SDCMS��վ��Ϣ����ϵͳ</td>
    </tr> 
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">�ٷ���վ��</td>
      <td><a href="http://www.sdcms.cn">www.sdcms.cn</a></td>
    </tr>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">��ǰ�汾��</td>
      <td><%=Sdcms_Version%></td>
    </tr>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
      <td height="22" align="center">����վ�㣺</td>
      <td><a href="http://www.nicemb.com" target="_blank">����ʱ������Ƽ����޹�˾</a></td>
    </tr>
  </table>
</div>

</body>
</html>
<%End Sub%>