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
	Case "add","edit","save","del","pass","nice","top","makehtml":title="��������"
	Case Else:stype="main":title="��Ϣ����"
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

<div class="sdcms_notice"><span>���������</span><a href="?action=add">�����Ϣ</a>������<a href="?">��Ϣ����</a>������<a href="?t=1">Ͷ�����</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><a<%if stype<>"main" then%> href="javascript:void(0)" onClick="selectTag('tagContent0',this)"<%end if%>><%=title%></a></li>
	<%if stype<>"main" then%>
	<li class="unsub"><a href="javascript:void(0)" onClick="selectTag('tagContent1',this)">��������</a></li>
	<%end if%>
	<%if stype="main" then%><li class="unsub"><a class="hand" onclick="$('#search')[0].style.visibility='inherit'">��Ϣ����</a></li><%end if%>
</ul><div style="visibility:hidden;position:absolute;margin:0 0 0 88px;*margin:0 0 0 -89px;border:1px solid #FCBA72;background:#fff;padding:5px 10px;width:280px;" id="search"><img src="images/close.gif"  style="position:absolute;margin:10px 0 0 242px;cursor:pointer;" onclick="$('#search')[0].style.visibility='hidden'" alt="�ر�" /><form action="?t=<%=t%>">�ؼ��֣�<input name="keyword" class="input" value="<%=keyword%>" /> <input type="submit" class="bnt01" value="����"></form>
</div>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "add":sdcms.Check_lever 12:add
	Case "edit":sdcms.Check_lever 13:add
	Case "save":save
	Case "pass":sdcms.Check_lever 13:pass(1)
	Case "nopass":sdcms.Check_lever 13:pass(0)
	Case "nice":sdcms.Check_lever 13:nice(1)
	Case "nonice":sdcms.Check_lever 13:nice(0)
	Case "top":sdcms.Check_lever 13:top(1)
	Case "notop":sdcms.Check_lever 13:top(0)
	Case "makehtml":sdcms.Check_lever 13:makehtml
	Case "movelist":sdcms.Check_lever 13:movelist
	Case "del":sdcms.Check_lever 14:del
	
	Case "go":Make_Class_Arr
	Case "pagelist":Make_Class_Page
	
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
      <td class="title_bg">����</td>
      <td width="40" class="title_bg">����</td>
      <td width="60" class="title_bg">����</td>
	  <td width="40" class="title_bg">����</td>
	  <td width="40" class="title_bg">��ͼ</td>
	  <td width="100" class="title_bg">���</td>
	  <td width="140" class="title_bg">״̬</td>
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
	.Table="View_Info"
	.Field="id,title,comment_num,hits,isurl,ispic,ispass,isnice,ontop,classid,ClassName,style,ClassUrl,HtmlName"
	.Key="ID"
	.Where=Where
	.Order="ontop desc,id desc"
	.PageStart="?classid="&classid&"&keyword="&keyword&"&t="&t&"&page="
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
	%>
    <tr onmouseover=this.bgColor='#EEFEED'; onmouseout=this.bgColor='#ffffff';  bgcolor='#ffffff'>
	  <td height="25" align="center"><input name="id" type="checkbox" value="<%=Rs(0)%>"></td>
      <td <%=Rs(11)%>><a href="<%=Get_Link(Sd_Table,Url)%>" target="_blank"><%=Rs(1)%></a></td>
	  <td align="center"><a href="Sdcms_Comment.Asp?Classid=<%=Rs(0)%>"><%=Rs(2)%></a></td>
	  <td align="center"><%=rs(3)%></td>
	  <td align="center"><%=IIF(Rs(4)=1,"��","<span class=""c9"">��</span>")%></td>
	  <td align="center"><%=IIF(Rs(5)=1,"��","<span class=""c9"">��</span>")%></td>
	  <td align="center"><a href="?classid=<%=Rs(9)%>"><%=Rs(10)%></a></td>
	  <td align="center"><%=IIF(Rs(6)=1,"����֤","<span class=""c9"">δ��֤</span>")%>&nbsp;<%=IIF(Rs(7)=1,"���Ƽ�","<span class=""c9"">δ�Ƽ�</span>")%>&nbsp;<%=IIF(Rs(8)=1,"���ö�","<span class=""c9"">δ�ö�</span>")%></td>
      <td align="center"><%IF Sdcms_Mode=2 Then%><a href="?action=makehtml&id=<%=rs(0)%>&t=<%=t%>&Page=<%=Page%>">����</a> <%End IF%><a href="?action=edit&id=<%=rs(0)%>&t=<%=t%>&Page=<%=Page%>">�༭</a> <a href="?action=del&id=<%=rs(0)%>&t=<%=t%>&Page=<%=Page%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
    </tr>
	<%
		Rs.MoveNext
	Next       
	%>
	<tr>
      <td colspan="9" class="tdbg">
	  <span class="right"><select onchange="location.href='?classid='+this.value+'&page=<%=page%>&t=<%=t%>'"><option value="0">���з���</option><%Echo Get_Class(Classid):PrintClass=Empty%></select></span>��<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox"><label for="chkall">ȫѡ</label>  
              <select name="action" onchange="if(this.value=='movelist'){t0.disabled=false}else{t0.disabled=true};">
			  <option>������</option>
			  <option value="pass">ͨ����֤</option>
			  <option value="nice">��Ϊ�Ƽ�</option>
			  <option value="top">��Ϊ�ö�</option>
			  <optgroup></optgroup>
			  <option value="nopass">ȡ����֤</option>
			  <option value="nonice">ȡ���Ƽ�</option>
			  <option value="notop">ȡ���ö�</option>
			  <optgroup></optgroup>
			  <option value="movelist">ת�Ʒ���</option>
			  <%IF Sdcms_Mode=2 Then%><option value="makehtml">������Ϣ</option><%End IF%>
			  <option value="del">ɾ����Ϣ</option>
			  </select>  <select name='t0' disabled="disabled"><option value="0">��ѡ�����</option><%=Get_Class(0)%></select>
      <input type="submit" class="bnt01" value="ִ��">

</td>
    </tr>
	<%IF Len(Num)=0 Then%>
	<tr>
      <td colspan="9" class="tdbg content_page" align="center"><%Echo P.PageList%></td>
	</tr>
	<%End IF%>
	</form>
  </table>

<%
Set P=Nothing
End Sub

Sub Add
	Dim Sql,Rs,Rs1
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID>0 Then
		
		Sql="Select Title,Pic,Classid,Topic,Isnice,Ontop,Ispass,Iscomment,Tags,Isurl,Url,Content,"'0-11
		Sql=Sql&"Author,ComeFrom,Hits,Htmlname,Keyword,JJ,LikeIdType,LikeID,adddate,Style From "&Sd_Table&" Where ID="&ID'12-21
		Set Rs=Conn.Execute(Sql)
		DbQuery=DbQuery+1
		IF Rs.Eof Then
			Echo "����Ƿ��ύ����":Exit Sub
		Else
			Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11
			Dim t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,Color,Pic_List
			
			t0=Rs(0)
			t1=Rs(1)
			t2=Rs(2)
			t3=Rs(3)
			t4=Rs(4)
			t5=Rs(5)
			t6=Rs(6)
			t7=Rs(7)
			t8=Rs(8)
			t9=Rs(9)
			t10=Rs(10)
			t11=Rs(11)
			t12=Rs(12)
			t13=Rs(13)
			t14=Rs(14)
			t15=Rs(15)
			t16=Rs(16)
			t17=Rs(17)
			t18=Rs(18)
			t19=Rs(19)
			t20=Rs(20)
			t21=Rs(21)
			IF Load_Cookies("sdcms_admin")=0 Then
				IF Instr(sdcms_infolever,t2&"|2")=0 Then Echo "��û�д���Ŀ��Ϣ�ı༭Ȩ��":Exit Sub
			End IF
			IF InStr(t21,"Color:")>0 Then Color=Mid(t21,14,7)
			IF Len(Color)=0 Then Color="#D4D0C8"
			Pic_List=Get_ImgSrc(t11)
		End IF
	Else
		t2=0
		t6=1
		t7=t6
		t9=t2
		t12=LoadRecord("PenName","Sd_Admin",sdcms_adminid)
		t14=t2
		t15=Sdcms_FileName
		t20=Dateadd("h",Sdcms_TimeZone,Now())
		Color="#D4D0C8"
	End IF
	DbQuery=DbQuery+1
	Echo Check_Add
%>
  <form name="add" method="post" action="?action=Save&ID=<%=ID%>&Page=<%=Page%>&t=<%=t%>" onSubmit="return checkadd()">
    <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" id="tagContent0">
	<tr>
      <td width="120" align="center" class="tdbg">��Ϣ���⣺</td>
      <td class="tdbg"><input name="t0" type="text" value="<%=t0%>" class="input" id="t0" size="50">
	  <input type="checkbox" name="c0_0" id="c0_0" value="font-weight:bold;" <%=IIF(Instr(t21,"bold")>0,"checked","")%> /><label for="c0_0">����</label>
	  <input type="checkbox" name="c0_1" id="c0_1" value="font-style:italic;" <%=IIF(Instr(t21,"italic")>0,"checked","")%> /><label for="c0_1">б��</label>
      <input type="hidden" id="c0_2" name="c0_2" value="<%=IIF(Color="#D4D0C8","",Color)%>" /><img alt="������ɫ" align="absmiddle" src="Images/color_selecter.gif" class="hand" id="color" style="background:<%=Color%>" /> <a href="javascript:void(0);" onclick="$('#color').css('background-color','#D4D0C8');$('#c0_2').val('');alert('����ɹ�')"><span>�����ɫ</span></a></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">�������ã�</td>
      <td class="tdbg"><select name="t2"><option value="">��ѡ�����</option><%=Get_Class(t2)%></select>
	  <select name="t3">
	  <option value="0" <%=IIF(t3=0,"selected","")%>>������ר��</option>
	  <%Set Rs1=Conn.Execute("Select id,title From Sd_Topic Order by Ordnum Desc"):While Not Rs1.Eof%>
	  <option value="<%=Rs1(0)%>" <%=IIF(Rs1(0)=t3,"selected","")%>><%=Rs1(1)%></option>
	  <%Rs1.Movenext:Wend:Rs1.Close%>
	  </select>
	  <input name="t4" id="t4" type="checkbox" value="1" <%=IIF(t4=1,"checked","")%> /><label for="t4">�Ƽ�</label>
	  <input name="t5" id="t5" type="checkbox" value="1" <%=IIF(t5=1,"checked","")%> /><label for="t5">�ö�</label>
	  <input name="t6" id="t6" type="checkbox" value="1" <%=IIF(t6=1,"checked","")%> /><label for="t6">��֤</label>
	  <input name="t7" id="t7" type="checkbox" value="1" <%=IIF(t7=1,"checked","")%> /><label for="t7">��������</label>
	  </td>
    </tr>
	<tr>
      <td align="center" class="tdbg">��ǩ������</td>
      <td  class="tdbg"><input name="t8" type="text" value="<%=t8%>" class="input" size="50" maxlength="250">��<span>֧�ֿո񡢶��ŷָ�</span></td>
    </tr>
    <tr class="tdbg">
      <td align="center">�ⲿ���ӣ�</td>
      <td><input name="t9" type="radio" onClick=$("#flag1")[0].style.display='none';$("#flag2")[0].style.display='inline';$("#uploadList")[0].style.display='inline';this.form.t10.disabled=true; value="0" <%=IIF(t9=0,"checked","")%> id="t9_0"><label for="t9_0">��</label>
	  <input name="t9" type="radio" onClick=$("#flag1")[0].style.display='inline';$('#flag2')[0].style.display='none';$("#uploadList")[0].style.display='none';this.form.t10.disabled=false; value="1" <%=IIF(t9=1,"checked","")%> id="t9_1"><label for="t9_1">��</label></td>
    </tr>
    <tr class="tdbg<%=IIF(t9=0," dis","")%>" id="flag1">
      <td align="center">���ӵ�ַ��</td>
      <td><input name="t10" type="text" value="<%=t10%>" <%=IIF(t9=0,"disabled","")%> class="input" id="t10" size="40">��<span>����д����·�� �磺http://www.sdcms.cn</span></td>
   </tr>
   <tr class="tdbg">
      <td align="center">����ժҪ��</td>
      <td valign="top"><input type="checkbox" value="1" onclick='Display("sdcms_intro")' name="intro" id="intro" /><label for="intro">�༭����ժҪ</label></td>
    </tr>
    <tr class="tdbg" id="sdcms_intro" style="display:none">
      <td align="center"></td>
      <td valign="top"><textarea name="O6" id="O6" cols="60" rows="4" style="width:100%;height:120px;" ><%=Content_Encode(t17)%></textarea></td>
    </tr>
   <tr class="tdbg<%=IIF(t9=1," dis","")%>" id="flag2">
      <td align="center">��Ϣ���ݣ�</td>
      <td><textarea name="t11" id="t11" style="width:100%;height:300px;"><%=Content_Encode(t11)%></textarea>
	  <input name="up" id="up" type="checkbox" value="1" /><label for="up">����Զ��ͼƬ</label><input type="checkbox" value="1" name="up1" id="up1" /><label for="up1">��ȡ�����е�һ��ͼƬΪ����ͼ</label>
	  </td>
    </tr>
	<tr class="tdbg" id="sdcms_pic">
      <td align="center">�� �� ͼ��</td>
      <td><input name="t1" id="t1" type="text" value="<%=t1%>" class="input" size="40"> <select id="uploadList" style="width:300px;display:none;" onchange="$('#t1')[0].value=this.value"><option value="<%=t1%>">��ѡ��</option><%=Pic_List%></select><br>
	  <%admin_upfile 1,"100%","20","t1","UpLoadPicIframe",1,1%></td>
    </tr>
    <%IF Sdcms_Mode=2 Then%>
    <tr class="tdbg">
      <td align="center">����ѡ�</td>
      <td>
     <input name="h0" type="checkbox" value="1" id="h0" <%=IIF(Instr(", "&Sdcms_Create_Set&", ",", 0, ")>0,"Checked","")%> /><label for="h0">������ҳ</label>
     <input name="h1" type="checkbox" value="1" id="h1" <%=IIF(Instr(", "&Sdcms_Create_Set&", ",", 1, ")>0,"Checked","")%> /><label for="h1">���ɷ���</label>
     <input name="h2" type="checkbox" value="1" id="h2" <%=IIF(Instr(", "&Sdcms_Create_Set&", ",", 2, ")>0,"Checked","")%> /><label for="h2">������Ϣ</label>
     </td>
    </tr>
    <%End IF%>
  </table>
   <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" id="tagContent1" class="dis">
	<tr class="tdbg">
      <td align="center" width="120">�����ߣ�</td>
      <td><input name="O1" value="<%=t12%>" type="text" class="input" size="40"></td>
    </tr>
	<tr class="tdbg">
      <td align="center">����Դ��</td>
      <td><input name="O2" id="O2" value="<%=t13%>" type="text" class="input" size="40">
	  <select onchange="$('#O2').val(this.value)">
		  <option value="">ѡ��</option>
		  <option value="δ֪">δ֪</option>
		  <option value="ԭ��">ԭ��</option>
		  <option value="ת��">ת��</option>
	  </select></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ˡ�����</td>
      <td><input name="O3" type="text" value="<%=t14%>" maxlength="6" onKeypress="if(event.keyCode<45||event.keyCode>57)event.returnValue=false;" class="input" size="40"> <span>���Զ������֣���߲�������999999��</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ļ�����</td>
      <td><input name="O4" value="<%=t15%>" type="text" class="input" size="40" /><%=Sdcms_FileTxt%></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ؼ��֣�</td>
      <td><input name="O5"  type="text" class="input" value="<%=t16%>" size="50"></td>
    </tr>
	
	<tr class="tdbg">
      <td align="center">������£�</td>
      <td><input name="O7" type="radio" value="0" <%=IIF(t18=0,"checked","")%> id="t20_1" onclick="$('#likeid')[0].style.display='none';" /><label for="t20_1">��ǩ���</label><input name="O7" type="radio" value="1" <%=IIF(t18=1,"checked","")%> id="t20_2" onclick="$('#likeid')[0].style.display='block';" /><label for="t20_2">ָ��ID</label></td>
    </tr>
	<tr class="tdbg<%=IIF(t18=0," dis","")%>" id="likeid">
      <td align="center">ָ��ID�� </td>
      <td valign="top"><textarea name="O8" cols="60" rows="2" class="inputs"><%=t19%></textarea><span>ID֮������","(Ӣ�Ķ���)��</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">¼��ʱ�䣺</td>
      <td><input name="O9" type="text" class="input" size="40" value="<%=t20%>" ></td>
    </tr>
</table>

<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" >
    <tr class="tdbg">
	  <td width="100">&nbsp;</td>
      <td><input type="submit" class="bnt" value="��������"> <input type="button" onClick="history.go(-1)" class="bnt" value="��������"></td>
    </tr>
</table>
 </form>
<%
	IF ID>0 Then
		Rs.Close
		Set Rs=Nothing
	End IF
End Sub

Sub Save
	Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11
	Dim O0,O1,O2,O3,O4,O5,O6,O7,O8,O9
	Dim Up,Up1,c0_0,c0_1,c0_2,Style,IsPic,ID
	Dim H0,H1,H2
	ID=IsNum(Trim(Request.QueryString("ID")),0)
	t0=FilterText(Trim(Request.Form("t0")),0)
	t1=FilterText(Trim(Request.Form("t1")),0)
	t2=FilterText(Trim(Request.Form("t2")),0)
	t3=FilterText(Trim(Request.Form("t3")),0)
	t4=IsNum(Trim(Request.Form("t4")),0)
	t5=IsNum(Trim(Request.Form("t5")),0)
	t6=IsNum(Trim(Request.Form("t6")),0)
	t7=IsNum(Trim(Request.Form("t7")),0)
	t8=Trim(Request.Form("t8"))
	t9=IsNum(Trim(Request.Form("t9")),0)
	t10=FilterText(Trim(Request.Form("t10")),0)
	t11=Request.Form("t11")
	
	O0=FilterText(Trim(Request.Form("O0")),0)
	O1=FilterText(Trim(Request.Form("O1")),0)
	O2=FilterText(Trim(Request.Form("O2")),0)
	O3=IsNum(Trim(Request.Form("O3")),0)
	O4=FilterText(Trim(Request.Form("O4")),0)
	O5=FilterHtml(Trim(Request.Form("O5")))
	O6=FilterHtml(Trim(Request.Form("O6")))
	O7=IsNum(Trim(Request.Form("O7")),0)
	O8=FilterHtml(Trim(Request.Form("O8")))
	O9=Trim(Request.Form("O9"))
	
	Up=IsNum(Trim(Request.Form("Up")),0)
	Up1=IsNum(Trim(Request.Form("Up1")),0)
	c0_0=FilterHtml(Trim(Request.Form("c0_0")))
	c0_1=FilterHtml(Trim(Request.Form("c0_1")))
	c0_2=FilterHtml(Trim(Request.Form("c0_2")))
	
	h0=IsNum(Trim(Request.Form("h0")),0)
	h1=IsNum(Trim(Request.Form("h1")),0)
	h2=IsNum(Trim(Request.Form("h2")),0)
	
	Style="Style="""
	IF Len(c0_2)>0 Then Style=Style&"Color:"&c0_2&";"
	IF Len(c0_0)>0 Then Style=Style&c0_0
	IF Len(c0_1)>0 Then Style=Style&c0_1
	
	Style=Style&""""
	IF Style="Style=""""" Then Style=""
	
	IF Len(O9)=0 Or Not IsDate(O9) Then O9=Dateadd("h",Sdcms_TimeZone,Now())
	t8=Re(t8," ",","):t8=Re(t8,"��",","):t8=Re(t8,"��",","):t8=Re(t8,"��",",")
	t8=FilterHtml(t8)
	IF O7=0 Then
		Dim Sdcms_Label,LikeID,LikeIDtag,LikeIDval,LikeIDData,I,Rs,Sql
		Sdcms_Label=Split(t8,",")
		LikeID=Empty
  		For I=0 To Ubound(Sdcms_Label)
			Set Rs=Conn.Execute("Select Top 10 ID From "&Sd_Table&" Where Title Like '%"&Sdcms_Label(I)&"%' Or Tags Like '%"&Sdcms_Label(I)&"%' ")
			While Not Rs.Eof
				LikeID=LikeID&Rs(0)&","
			Rs.MoveNext
			Wend
			Rs.Close
		Next
		IF Len(LikeID)>0 Then O8=LikeID Else O8=0
	End IF
	O8=Check_Event(O8,",","")
	If Len(O8)=0 Then O8=0

	IF Up=1 Then t11=ReplaceRemoteUrl(t11,"","","","")
	IF Up1=1 Then t1=Frist_Pic(t11)
	IsPic=Check_ispic(t1)
	
	IF ID=0 Then
		IF t2>0 And Load_Cookies("sdcms_admin")=0 Then
			IF Instr(sdcms_infolever,t2&"|1")=0 Then Echo "��û�д���Ŀ�Ĵ���Ȩ��":Exit Sub
		End IF
	Else
		IF Load_Cookies("sdcms_admin")=0 Then
			IF Instr(sdcms_infolever,t2&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Exit Sub
		End IF
	End IF
	
	IF Id>0 Then
		Dim Old_ClassUrl
		Set Rs=Conn.Execute("Select ClassUrl From View_Info Where ID="&ID&"")
		IF Not Rs.Eof Then
			Old_ClassUrl=Rs(0)
		End IF
	End IF
	
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select Title,Pic,Classid,Topic,Isnice,Ontop,Ispass,Iscomment,Tags,Isurl,Url,Content,"'0-11
	Sql=Sql&"Author,ComeFrom,Hits,Htmlname,Keyword,JJ,LikeIdType,LikeID,adddate,Style,IsPic,UserID,LastUpDate,ID From "&Sd_Table&" "'12-25
	IF ID>0 Then 
		Sql=Sql&" Where ID="&ID
	End IF
	Rs.Open Sql,Conn,1,3
	IF ID=0 Then 
		Rs.Addnew
	Else
		Rs.Update
	End IF
	Rs(0)=Left(t0,255)
	Rs(1)=Left(t1,255)
	Rs(2)=t2
	Rs(3)=t3
	Rs(4)=t4
	Rs(5)=t5
	Rs(6)=t6
	Rs(7)=t7
	IF ID>0 Then Lost_tags(Rs(8))'�ȼ�ȥԭ���Ĵ���
	Add_tags(t8) '������ʹ�ô���
	Rs(8)=Left(t8,255)
	Rs(9)=t9
	Rs(10)=Left(t10,50)
	Rs(11)=t11
	Rs(12)=Left(O1,50)
	Rs(13)=Left(O2,50)
	Rs(14)=Left(O3,50)
	'IF Sdcms_Mode=2 And Id>0 And t=0 Then
		'Del_File Sdcms_Root&Sdcms_HtmDir&Old_ClassUrl&Rs(15)&Sdcms_FileTxt
		'Dim ReCreate:ReCreate=1
	'End IF
	Rs(15)=O4
	Rs(16)=Left(O5,50)
	IF Len(O6)=0 Or IsNull(O6) Then
		Rs(17)=CloseHtml(CutStr(Content_Decode(Re_Html(t11)),Sdcms_Length))
	Else
		Rs(17)=CloseHtml(Content_Decode(O6))
	End IF
	Rs(18)=O7
	Rs(19)=O8
	Rs(20)=O9
	Rs(21)=Style
	Rs(22)=IsPic
	IF ID>0 And Clng(t6)=1 And Rs(23)<0 Then Rs(23)=0
	Rs(24)=Now()
	
	Rs.UpDate
	Rs.MoveLast
	ID=Rs(25)
	'������������ǩ,��Ҫ�滻��
	Custom_HtmlName O4,sd_table,t0,ID
	Rs.Close
	Set Rs=Nothing
	'IF Len(ReCreate)>0 Then
		'Set Sdcms=New Sdcms_Create
		  'Sdcms.Create_Info_Show ID
		'Set Sdcms=Nothing
	'End IF
	IF Sdcms_Mode=2 Then
		IF Clng(h0)+Clng(h1)+Clng(h2)>0 Then
			Dim Sdcms_C
			Set Sdcms_C=New Sdcms_Create
			IF h0>0 Then
				Sdcms_C.Create_Index
			End IF
			IF h2>0 Then
				Sdcms_C.Create_Info_Show ID
			End IF
			Echo "<br>ȫ���������"
			IF h1>0 Then
				'��ȡ��ID������ParenetID
				Dim PartentID
				PartentID=Conn.Execute("Select PartentID From Sd_Class Where ID="&t2&"")(0)
				Dim Parray
				Parray=Split(PartentID,",")
				Dim Total_Class_Num
				Total_Class_Num=Ubound(Parray)+1
				Add_Cookies "ClassIDArray",PartentID
				Go "?action=go&Total_Class_Num="&Total_Class_Num&"&This_Arr=1"
				'Set Rs=Conn.Execute("Select AllClassID,PageNum,Class_Type From Sd_Class Where ID="&t2&"")
				'IF Rs.Eof Then
					'Echo "��������"
				'Else
					'IF Rs(2)=1 Then
						'Sdcms_C.Create_Channel ID
					'Else
						'Dim This_Count,TotalPage
						'This_Count=Conn.Execute("Select Count(ID) From "&Sd_Table&" Where IsPass=1 And ClassID In("&Rs(0)&")")(0)
						'TotalPage=Abs(Int(-Abs(This_Count/Rs(1))))
						'IF TotalPage=0 Then TotalPage=1
						'Go "?action=pagelist&id="&t2&"&TotalPage="&TotalPage&"&page=1"
					'End IF
				'End IF
				'Rs.Close
				'Set Rs=Nothing
			End IF
			Set Sdcms_C=Nothing
		Else
			Go "?t="&t&"&Page="&Page
		End IF
	Else
		Go "?t="&t&"&Page="&Page
	End IF
End Sub

Sub Make_Class_Arr
	Dim Total_Class_Num,This_Arr
	Total_Class_Num=IsNum(Trim(Request.QueryString("Total_Class_Num")),0)
	This_Arr=IsNum(Trim(Request.QueryString("This_Arr")),0)
	Echo "�ܼ���Ҫ���ɣ�<b>"&Total_Class_Num&" </b>����Ŀ�������ɣ�<b>"&This_Arr&"</b> ��<br><br>"
	
	Dim This_ID,Class_Arr
	Class_Arr=Load_Cookies("ClassIDArray")
	Class_Arr=Split(Class_Arr,",")
	
	This_ID=Class_Arr(This_Arr-1)
	'======================================
	Dim Rs
	Set Rs=Conn.Execute("Select AllClassID,PageNum,Class_Type From Sd_Class Where ID="&This_ID&"")
	IF Not Rs.Eof Then
		IF Rs(2)=1 Then
			Dim Sdcms_C
			Set Sdcms_C=New Sdcms_Create
			Sdcms_C.Create_Channel This_ID
			Set Sdcms_C=Nothing
		Else
			Dim This_Count,TotalPage
			This_Count=Conn.Execute("Select Count(ID) From Sd_Info Where IsPass=1 And ClassID In("&Rs(0)&")")(0)
			TotalPage=Abs(Int(-Abs(This_Count/Rs(1))))
			IF TotalPage=0 Then TotalPage=1
			Go "?action=pagelist&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"&id="&This_ID&"&TotalPage="&TotalPage&"&page=1":Died
		End IF
	End IF
	Rs.Close
	Set Rs=Nothing
	'======================================
	This_Arr=This_Arr+1
	
	IF This_Arr<=Total_Class_Num Then
		Echo "<script>setTimeout(""location.href='?act=go&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"';"",""100"");</script>"
	Else
		Echo "<br><b>ȫ���������</b><br>"
	End IF
End Sub

Sub Make_Class_Page
	Dim Total_Class_Num,This_Arr
	Total_Class_Num=IsNum(Trim(Request.QueryString("Total_Class_Num")),0)
	This_Arr=IsNum(Trim(Request.QueryString("This_Arr")),0)
	
	Echo "�ܼ���Ҫ���ɣ�<b>"&Total_Class_Num&"</b> ����Ŀ�������ɣ�<b>"&This_Arr&"</b> ��<br>"
	
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	Dim TotalPage:TotalPage=IsNum(Trim(Request.QueryString("TotalPage")),0)
	Dim Pages:Pages=IsNum(Trim(Request.QueryString("Page")),0)
	Echo "<br>�ܼ���Ҫ���ɣ�"&TotalPage&" ҳ �����ɣ�"&Pages&" ҳ<br><br>"
	Dim Sdcms_C
	Set Sdcms_C=New Sdcms_Create
	Sdcms_C.Create_I_List ID
	Set Sdcms_C=Nothing
	Pages=Pages+1
	
	IF Pages<=TotalPage Then
		Echo "<script>setTimeout(""location.href='?action=pagelist&id="&id&"&TotalPage="&TotalPage&"&page="&Pages&"&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"';"",""100"");</script>"
	Else
		IF This_Arr>=Total_Class_Num Then
			Echo "<br><b>�������</b><br>":Exit Sub
		End IF
		Echo "<script>setTimeout(""location.href='?action=go&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr+1&"';"",""100"");</script>"
	End IF

End Sub

Sub MakeHtml
	Dim ID:ID=FilterHtml(Trim(Request("ID")))
	Dim I
	AddLog sdcms_adminname,GetIp,"������Ϣ�����Ϊ"&ID,0
	ID=Split(ID,", ")
	For I=0 To Ubound(ID)
	  Set Sdcms=New Sdcms_Create
		  Sdcms.Create_Info_Show Clng(ID(I))
		  Response.Flush()
	  Set Sdcms=Nothing
	Next
	Echo "<br>�������"
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	IF Len(ID)>0 then
		AddLog sdcms_adminname,GetIp,"ɾ����Ϣ�����Ϊ"&id,0
		ID=Split(ID,", ")
		Dim I,ClassUrl,HtmlName,Tags
		For I=0 To Ubound(ID)			
			Set Rs=Conn.Execute("Select ClassUrl,HtmlName,Tags,ClassID From View_Info Where ID="&ID(I)&"")
			IF Not Rs.Eof Then
				ClassUrl=Rs(0)
				HtmlName=Rs(1)
				Tags=Rs(2)
				
				IF Load_Cookies("sdcms_admin")=0 Then
					IF Instr(sdcms_infolever,Rs(3)&"|3")=0 Then Echo "��û�д���Ŀ��ɾ��Ȩ��":Died
				End IF
				
				Lost_Tags Tags
				
				IF Sdcms_Mode=2 Then
					Del_File Sdcms_Root&sdcms_htmdir&ClassUrl&HtmlName&Sdcms_FileTxt
				End IF
				
				Conn.Execute("Delete From "&Sd_Table02&" Where infoid="&Clng(ID(I))&"")
				Conn.Execute("Delete From "&Sd_Table03&" Where Followid="&Clng(ID(I))&"")
				Conn.Execute("Delete From "&Sd_Table&" Where id="&Clng(ID(I))&"")
			End IF
		Next
	End IF
	Go "?t="&t&"&Page="&Page
End sub

Sub Pass(t0)
	Dim Msg,I,ID
	ID=FilterHtml(Trim(Request("ID")))
	Msg=IIF(t0=1,"�����Ϣ","ȡ�����")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,Msg&"�����Ϊ"&id,0
		ID=Split(ID,", ")
		For I=0 To Ubound(ID)
			IF Load_Cookies("sdcms_admin")=0 Then
				IF Instr(sdcms_infolever,Loadrecord("classid",Sd_Table,Clng(ID(I)))&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Died
			End IF
			Conn.Execute("Update "&Sd_Table&" Set IsPass="&t0&" Where ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "?t="&t&"&Page="&Page
End Sub

Sub Nice(t0)
	Dim Msg,I,ID
	ID=FilterHtml(Trim(Request("ID")))
	Msg=IIF(t0=1,"�Ƽ���Ϣ","ȡ���Ƽ�")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,Msg&"�����Ϊ"&id,0
		ID=Split(ID,", ")
		For I=0 To Ubound(ID)
			IF Load_Cookies("sdcms_admin")=0 Then
				IF Instr(sdcms_infolever,Loadrecord("classid",Sd_Table,Clng(ID(I)))&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Died
			End IF
			Conn.Execute("Update "&Sd_Table&" Set IsNice="&t0&" Where ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "?t="&t&"&Page="&Page
End Sub

Sub Top(t0)
	Dim Msg,I,ID
	ID=FilterHtml(Trim(Request("ID")))
	Msg=IIF(t0=1,"�ö���Ϣ","ȡ���ö�")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,Msg&"�����Ϊ"&id,0
		ID=Split(ID,", ")
		For I=0 To Ubound(ID)
			IF Load_Cookies("sdcms_admin")=0 Then
				IF Instr(sdcms_infolever,Loadrecord("classid",Sd_Table,Clng(ID(I)))&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Died
			End IF
			Conn.Execute("Update "&Sd_Table&" Set Ontop="&t0&" Where ID="&Clng(ID(I))&"")
		Next
	End IF
	Go "?t="&t&"&Page="&Page
End Sub

Sub Movelist
	Dim t1,i,t0,ids,rs,sql,id
	t1=t
	t0=IsNum(Request.Form("t0"),0)
	ID=Trim(Request.Form("ID"))
	IF Load_Cookies("sdcms_admin")=0 Then
	IF Instr(sdcms_infolever,t0&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Died
	End IF
	IF t0>0 Then
		IDS=Split(ID,", ")
		IF Len(ID)=0 Then Alert "����ѡ��һ����Ϣ","?":Died
		Set Sdcms=New Sdcms_Create
		For I=0 To Ubound(IDs)
			IF Load_Cookies("sdcms_admin")=0 Then
				IF Instr(sdcms_infolever,Loadrecord("classid",Sd_Table,Clng(ID(I)))&"|2")=0 Then Echo "��û�д���Ŀ�ı༭Ȩ��":Died
			End IF
			Set Rs=Server.CreateObject("adodb.recordset")
			Sql="Select Classid,htmlname,id,ClassUrl From View_Info Where Id="&Clng(IDs(I))&""
			Rs.Open Sql,conn,1,3
			IF Not Rs.Eof Then
				IF Sdcms_Mode=2 Then Del_File Sdcms_Root&Sdcms_HtmDir&Rs(3)&Rs(1)&sdcms_filetxt
				Rs.Update
				Rs(0)=t0
				Rs.Update
				IF Sdcms_Mode=2 Then sdcms.Create_info_show Rs(2):Response.Flush()
			End IF
			Rs.Close
			Set Rs=Nothing
		Next
		Set Sdcms=Nothing
		Alert "ת�Ƴɹ�","?t="&t1&"&Page="&Page
	Else
		Alert "��ѡ�����","?t="&t1&"&Page="&Page
	End IF	
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add& "KE.show({"
	Check_Add=Check_Add& "			id : 't11',"
	Check_Add=Check_Add& "			imageUploadJson : '../../../"&Get_ThisFolder&"Sdcms_Editor_Up.asp',"
	Check_Add=Check_Add& "			fileUploadJson : '../../"&Get_ThisFolder&"Sdcms_Editor_Up.asp?act=1'"
	Check_Add=Check_Add& "		});"
	Check_Add=Check_Add& "KE.show({"
	Check_Add=Check_Add& "			id : 'O6',"
	Check_Add=Check_Add& "			imageUploadJson : '../../../"&Get_ThisFolder&"Sdcms_Editor_Up.asp',"
	Check_Add=Check_Add& "			items : ["
	Check_Add=Check_Add& "				'fontname', 'fontsize', '|', 'textcolor', 'bgcolor', 'bold', 'italic', 'underline',"
	Check_Add=Check_Add& "				'|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist',"
	Check_Add=Check_Add& "				'insertunorderedlist', '|',  'image', 'link', 'unlink','source', 'about']"
	Check_Add=Check_Add& "		});"
	Check_Add=Check_Add& "$(function(){$.showcolor('color','c0_2');});"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{change_tab()"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��Ϣ���ⲻ��Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t2.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��ѡ�����');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t2.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (!document.add.t10.disabled && document.add.t10.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('���ӵ�ַ����Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t3.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t10.disabled && KE.isEmpty('t11'))"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('���ݲ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t11.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"��function change_tab()"&vbcrlf
	Check_Add=Check_Add&"��{"&vbcrlf
	Check_Add=Check_Add&"����$(""#tagContent0"")[0].style.display='block';"&vbcrlf
	Check_Add=Check_Add&"����$(""#tagContent1"")[0].style.display='none';"&vbcrlf
	Check_Add=Check_Add&"����$(""#sdcms_sub_title li"").removeClass();"&vbcrlf
	Check_Add=Check_Add&"����$(""#sdcms_sub_title li:first-child"").addClass(""sub"");"&vbcrlf
	Check_Add=Check_Add&"����$(""#sdcms_sub_title li:last-child"").addClass(""unsub"");"&vbcrlf
	Check_Add=Check_Add&"��}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>