<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,Sd_Table,title,Action
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=new Sdcms_Admin
Sdcms.Check_admin
Select Case Action
	Case "add":title="���ͶƱ"
	Case "edit":title="�޸�ͶƱ"
	Case "show":title="�鿴ͶƱ"
	Case Else:title="ͶƱ����"
End Select
Sd_Table="sd_vote"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="?action=add">���ͶƱ</a>������<a href="?">ͶƱ����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li>
	
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 21:Add
	Case "edit":sdcms.Check_lever 22:Add
	Case "save":save
	Case "del":sdcms.Check_lever 23:Del
	Case "show":shows
	Case Else:Main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing
Sub Main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b" id="tagContent0">
    <tr>
      <td width="60" class="title_bg">���</td>
      <td class="title_bg">ͶƱ����</td>
      <td width="160" class="title_bg">����</td>
      <td width="160" class="title_bg">����</td>
    </tr>
	<%
	Dim Rs,i
	Set Rs=Conn.Execute("select id,title,adddate  from "&Sd_Table&"   order by id desc")
	DbQuery=DbQuery+1
	While Not Rs.Eof
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	<%For I=0 To 2%>
      <td height="25" <%IF I<>1 then%>align="center"<%end if%>><%=Rs(I)%></td>
	  <%Next%>
      <td align="center"><a href='?action=show&id=<%=rs(0)%>'>�鿴</a> <a href="?action=edit&id=<%=Rs(0)%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%Rs.MoveNext:Wend:Rs.Close:Set Rs=Nothing%>
  </table>
  
<%
End Sub
Sub Add
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID>0 Then
		Dim Rs,i,vote,result
		Set Rs=Conn.Execute("select title,stype,vote,result from "&Sd_Table&" where id="&id&"")
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
		t1=1
	End IF
	Echo Check_Add
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=save&id=<%=id%>" onSubmit="return checkadd()">
    <tr>
      <td width="9%" align="center" class="tdbg">ͶƱ���ƣ�      </td>
      <td width="91%" class="tdbg"><input name="t0" type="text" class="input" value="<%=t0%>" id="t0" size="50"></td>
    </tr>
    <tr class="tdbg">
      <td align="center">ͶƱѡ�      </td>
      <td><input name="t1" type="radio" value="1" <%=IIF(t1=1,"checked","")%> id="t1_0"><label for="t1_0">��ѡ</label> <input name="t1" type="radio" value="0" <%=IIF(t1=0,"checked","")%> id="t1_1"><label for="t1_1">��ѡ</label></td>
    </tr>
    <tr class="tdbg">
      <td align="center">��Ŀ���ݣ�</td>
      <td><input type="button" class="bnt" name="addvote" value="�����Ŀ" onclick="Addvote();"> <input type="button" class="bnt" name="modifyvote" value="�޸���Ŀ" onclick="return Modifyvote();"> <input type="button" class="bnt" name="delvote" value="ɾ����Ŀ" onclick="Delvote();"> <input class="bnt" onClick="changepos(content,-1)"  value="�� ��" type="button"> <input class="bnt" onClick="changepos(content,1)" type="button" value="�� ��"><br><input type="hidden" name="votes" value="" ><select class="inputs" name="content" style="width:400px;height:200px;margin-top:5px;" size="2" ondblclick="return Modifyvote();" >
	<%
	IF Len(t3)>0 Then
		result=split(t3,"|")
		vote=split(t2,"|")
		For I=0 to Ubound(vote)-1
			Echo "<option value="&vote(i)&"|"&result(i)&">"&Content_Encode(vote(i))&"|"&Content_Encode(result(i))&"</option>"
		next
	End IF
	%>
</select></td>
    </tr>
	 
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��������"> <input type="button" onClick="history.go(-1)" class="bnt" value="��������"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Shows
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim rs,vote,result,total_vote,i
	Set Rs=Conn.Execute("select id,title,vote,result from sd_vote where id="&id&"")
	DbQuery=DbQuery+1
	IF Rs.Eof Then
		Go("?"):Exit Sub
	Else
		vote=rs(2)
		result=rs(3)
		total_vote=0
		vote=split(vote,"|")
		result=split(result,"|")
		For I=0 To Ubound(result)-1
			IF Not result(I)="" Then total_vote=result(i)+total_vote
		Next
	End IF
	Rs.Close
	Set Rs=Nothing
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="table_b">
     <tr>
      <td width="45%" align="center" class="title_bg" >ѡ��      </td>
      <td width="45%" class="title_bg" >���� </td>
      <td width="10%" class="title_bg" >Ʊ��</td>
    </tr>
	<%For I=0 To Ubound(vote)-1%>
    <tr class="tdbg">
	  <td height="25"><%Echo ""&i+1&". "&vote(i)&""%> (<%IF total_vote=0 Then Echo "0" Else Echo IIF(result(I)=0,"0",Formatpercent(result(I)/total_vote,0))%>)</td>
      <td><div style="border:1px solid #ccc;"><div style='width:<%IF total_vote=0 Then Echo "0" Else Echo IIF(result(I)=0,"0",Formatpercent(result(I)/total_vote,0))%>;background:#5BAA26;height:14px;'></div></div></td>
      <td align="center"><%=result(I)%></td>
    </tr>
	<%Next%>
	<tr>
      <td colspan="2" align="center" class="tdbg" ><input type="button" onClick="location.href='?action=edit&id=<%=ID%>'" class="bnt01" value="�༭">��<input type="button" onClick="history.go(-1)" class="bnt01" value="����"></td>
      <td align="center" class="tdbg" >������<%=total_vote%></td>
	</tr>

  </table>
<%
End Sub

Sub Save
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim t0,t1,votes,vote,result,i,somevote,j,s,s1,rs,sql,LogMsg
	t0=FilterText(Trim(Request.Form("t0")),1)
	t1=IsNum(Trim(Request.Form("t1")),0)
	votes=FilterHtml(Trim(Request.Form("votes")))
	IF ID=0 Then sdcms.Check_lever 21 Else sdcms.Check_lever 22
	IF Len(votes)=0 Then
		Alert "��Ŀ���ݲ���Ϊ��","javascript:history.go(-1)":Exit Sub
	End IF
	 
	votes=Split(votes,",")
	For i=0 To (Ubound(votes)-1)
		IF Not Instr(votes(i),"|")>0 Then
			Alert "ͶƱѡ��"&i+1&"�д���","javascript:history.go(-1)":Exit Sub
		End IF
		somevote=Split(votes(i),"|")
		For J=0 To Ubound(somevote)
		s=somevote(0)
		s1=somevote(1)
		IF Not Isnumeric(s1) Then
			Alert "ͶƱѡ��"&i+1&"�д���","javascript:history.go(-1)":Exit Sub
		End IF
		Next  
		vote=vote&s&"|"
		result=result&s1&"|"
	Next
	
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select id,vote,result,title,stype,adddate From "&Sd_Table&""
	IF ID>0 Then 
		Sql=Sql&" where id="&ID
	End IF
	Rs.Open Sql,Conn,1,3
	
	IF ID=0 Then 
		Rs.Addnew
	Else
		Rs.Update
	End IF
	rs(1)=vote
	rs(2)=result
	rs(3)=t0
	rs(4)=t1
	IF ID=0 Then Rs(5)=Dateadd("h",Sdcms_TimeZone,Now())
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	IF ID=0 Then LogMsg="���ͶƱ" Else LogMsg="�޸�ͶƱ"
	AddLog sdcms_adminname,GetIp,LogMsg&t0,0
	Go("?")
End Sub

Sub Del
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	AddLog sdcms_adminname,GetIp,"ɾ��ͶƱ�����Ϊ"&id,0
	Conn.Execute("Delete From "&Sd_Table&" Where id="&id&"")
	Go("?") 
End Sub

Function Check_Add
	Check_Add="<script>"&vbcrlf
	Check_Add=Check_Add&"function changepos(obj,index)"&vbcrlf
	Check_Add=Check_Add&"{"&vbcrlf
	Check_Add=Check_Add&"if(index==-1){"&vbcrlf
	Check_Add=Check_Add&"if (obj.selectedIndex>0){"&vbcrlf
	Check_Add=Check_Add&"obj.options(obj.selectedIndex).swapNode(obj.options(obj.selectedIndex-1))"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"else if(index==1){"&vbcrlf
	Check_Add=Check_Add&"if (obj.selectedIndex<obj.options.length-1){"&vbcrlf
	Check_Add=Check_Add&"obj.options(obj.selectedIndex).swapNode(obj.options(obj.selectedIndex+1))"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	
	Check_Add=Check_Add&"function Addvote(){"&vbcrlf
	Check_Add=Check_Add&"  var thisvote='ͶƱ����'+(document.add.content.length+1)+'|0'; "&vbcrlf
	Check_Add=Check_Add&"  var vote=prompt('������ͶƱ���ƺͳ�ʼֵ���м��á�|��������',thisvote);"&vbcrlf
	Check_Add=Check_Add&"  if(vote!=null&&vote!=''){document.add.content.options[document.add.content.length]=new Option(vote,vote);}"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"function Modifyvote(){"&vbcrlf
	Check_Add=Check_Add&"  if(document.add.content.length==0) return false;"&vbcrlf
	Check_Add=Check_Add&"  var thisvote=document.add.content.value; "&vbcrlf
	Check_Add=Check_Add&"  if (thisvote=='') {alert('����ѡ��һ��ͶƱ��Ŀ���ٵ��޸İ�ť��');return false;}"&vbcrlf
	Check_Add=Check_Add&"  var vote=prompt('������ͶƱ���ƺͳ�ʼֵ���м��á�|��������',thisvote);"&vbcrlf
	Check_Add=Check_Add&"  if(vote!=thisvote&&vote!=null&&vote!=''){document.add.content.options[document.add.content.selectedIndex]=new Option(vote,vote);}"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"function Delvote(){"&vbcrlf
	Check_Add=Check_Add&"  if(document.add.content.length==0) return false;"&vbcrlf
	Check_Add=Check_Add&"  var thisvote=document.add.content.value; "&vbcrlf
	Check_Add=Check_Add&"  if (thisvote=='') {alert('����ѡ��һ��ͶƱ��Ŀ���ٵ�ɾ����ť��');return false;}"&vbcrlf
	Check_Add=Check_Add&"  document.add.content.options[document.add.content.selectedIndex]=null;"&vbcrlf
	Check_Add=Check_Add&"}"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('ͶƱ���Ʋ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	 
	Check_Add=Check_Add&"	if (document.add.content.length==0)"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('ͶƱ��Ŀ����Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.content.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.content.length<2)"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('ͶƱ��Ŀ����������������');"&vbcrlf
	Check_Add=Check_Add&"	document.add.content.focus();"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"��var s=""""; "&vbcrlf
	Check_Add=Check_Add&"��for(i=0;i<=document.all(""content"").length-1;i++) "&vbcrlf
	Check_Add=Check_Add&"��{ "&vbcrlf
	Check_Add=Check_Add&"��var s=s+document.add.content.options(i).value+"",""; "&vbcrlf
	Check_Add=Check_Add&"��} "&vbcrlf
	Check_Add=Check_Add&"��document.add.votes.value=s"&vbcrlf
	Check_Add=Check_Add&"��}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>