<!--#include file="sdcms_check.asp"-->
<!--#include file="../Inc/AspJpeg.asp"-->
<!--#include file="../Plug/Coll_Info/Conn.asp"-->
<!--#include file="../Plug/Coll_Info/Function.asp"-->
<%
Dim sdcms,title,Sd_Table,Action
Action=Lcase(Trim(Request("Action")))
Set sdcms=New Sdcms_Admin
sdcms.Check_admin
Select Case action
	Case "add":title="�����Ŀ"
	Case "edit":title="�޸���Ŀ"
	Case "import","importnext":title="��������"
	Case "export":title="��������"
	Case Else:title="�ɼ�����"
End Select
Sd_Table="Sd_Coll_Item"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="Sdcms_Coll_Config.asp">�ɼ�����</a>������<a href="Sdcms_Coll_Item.asp">�ɼ�����</a> (<a href="Sdcms_Coll_Item.asp?action=add">���</a> | <a href="?action=Import">����</a> | <a href="?action=Export">����</a>)������<a href="Sdcms_Coll_Filters.asp">���˹���</a> (<a href="Sdcms_Coll_Filters.asp?action=add">���</a>)������<a href="Sdcms_Coll_History.asp">��ʷ��¼</a></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><%=title%></li>
</ul>
<div id="sdcms_right_b">
<%
Collection_Data
Select Case Action
	Case "add":sdcms.Check_lever 21:Step1
	Case "edit":sdcms.Check_lever 22:Step1
	Case "step1_save":Step1_Save
	Case "step2":Step2
	Case "step2_save":Step2_Save
	Case "step3":Step3
	Case "step3_save":Step3_Save
	Case "copy":sdcms.Check_lever 22:Copy
	Case "import":sdcms.Check_lever 22:Import
	Case "importnext":importnext
	Case "importover":importover
	Case "export":sdcms.Check_lever 22:Export
	Case "exportnext":Exportnext
	Case "del":sdcms.Check_lever 23:Del
	Case "pass":sdcms.Check_lever 22:Pass(1)
	Case "nopass":sdcms.Check_lever 22:Pass(0)
	Case Else:main
End Select
Db_Run
CloseDb
Set Sdcms=Nothing

Sub Main
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b" id="tagContent0">
    <form name="add" action="?" method="post"  onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<tr>
      <td width="30" class="title_bg">ѡ��</td>
      <td class="title_bg">��Ŀ����</td>
      <td width="100" class="title_bg">��������</td>
	  <td width="100" class="title_bg">����ר��</td>
	  <td width="60" class="title_bg">״̬</td>
	  <td width="140" class="title_bg">�ϴβɼ�</td>
      <td width="220" class="title_bg">����</td>
    </tr>
	<%
	Dim Page,P,Rs,i,num,rs1
	Page=IsNum(Trim(Request.QueryString("page")),1)
	Set P=New Sdcms_Page
	With P
	.Conn=Coll_Conn
	.PageNum=Page
	.Table=Sd_Table
	.Field="id,ItemName,ClassID,SpecialID,Flag,UpDateTime"
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
	  <td align="center"><%IF Rs(2)=0 Then%>δָ��<%Else%><%Set Rs1=Conn.Execute("Select title From Sd_Class Where Id="&Clng(Rs(2))&""):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "<b>��������</b>":End IF%><%End IF%></td>
	  <td align="center"><%IF Rs(3)=0 Then%>δָ��<%Else%><%Set Rs1=Conn.Execute("Select title From Sd_Topic Where Id="&Clng(Rs(3))&""):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "��������":End IF%><%End IF%></td>
	  <td align="center"><%=IIF(Rs(4),"��","<b>��</b>")%></td>
	  <td align="center"><%Set Rs1=Coll_Conn.Execute("Select Adddate From Sd_Coll_History Where ItemID="&Clng(Rs(0))&" And Result=1 Order By Id Desc"):IF Not Rs1.Eof Then Echo Rs1(0):Else Echo "û�м�¼":End IF%></td>
      <td align="center"><a href="?action=Copy&id=<%=rs(0)%>" onclick='return confirm("ȷ��Ҫ���ƣ�");'>����</a> <%IF Rs(4)=1 Then%><a href="?action=NoPass&id=<%=rs(0)%>">����</a><%Else%><a href="?action=Pass&id=<%=rs(0)%>">����</a><%End IF%> <a href="?action=edit&id=<%=rs(0)%>">�༭</a> <a href="Sdcms_Coll_Coll.asp?action=Coll&id=<%=rs(0)%>">�ɼ�</a> <a href="Sdcms_Coll_Coll.asp?action=Demo&id=<%=rs(0)%>">����</a> <a href="?action=del&id=<%=rs(0)%>" onclick='return confirm("���Ҫɾ��?���ɻָ�!");'>ɾ��</a></td>
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

Sub step1
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,t24,t25,t26,t27,t28,t29,t30,t31,t32,t33,Sql,Rs,I
Check_Info
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
Sql="Select top 1 ItemName,Classid,selEncoding,ListStr,ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,LPsString,LPoString"
Sql=Sql&",Passed,SaveFiles,Thumb_WaterMark,Flag,CollecOrder,Coll_Top,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class"
Sql=Sql&",Script_table,Script_tr,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,Script_Td,strReplace,CollecNewsNum,Hits,SpecialID From "&Sd_Table&" Where Id="&ID
Set Rs=Coll_Conn.Execute(Sql)
IF Not Rs.Eof Then
	t0=Rs(0)
	t1=Rs(1)
	t2=Rs(2)
	t3=Rs(3)
	t4=Rs(4)
	t5=Rs(5)
	t6=Rs(6)
	t7=Rs(7)
	t8=Rs(8)
	't9=Rs(9)
	't10=Rs(10)
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
	t22=Rs(22)
	t23=Rs(23)
	t24=Rs(24)
	t25=Rs(25)
	t26=Rs(26)
	t27=Rs(27)
	t28=Rs(28)
	t29=Rs(29)
	t30=Rs(30)
	t31=Rs(31)
	t32=Rs(32)
	t33=Rs(33)
Else
	t11=1
	t14=1
	t15=1
End IF
t1=IsNum(t1,0)
t4=IsNum(t4,0)
t16=IsNum(t16,0)
t31=IsNum(t31,0)
t32=IsNum(t32,0)
t33=IsNum(t33,0)
IF Len(t8)>0 Then
	t8=Re(t8,"|","")
End IF
Echo "��Ŀ���ã�<a href=""?action="&action&"&id="&id&"""><b>��һ��</b></a>"
IF Id>0 Then
	Echo " >> <a href=""?action=step2&id="&id&""">�ڶ���</a> >> <a href=""?action=step3&id="&id&""">������</a>"
Else
	Echo " >> �ڶ��� >> ������"
End IF
Echo "<br><br>"
%>
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=step1_save&id=<%=id%>" onSubmit="return checkadd()">
    <tr>
      <td width="120" align="right" class="tdbg">��Ŀ���ƣ�</td>
      <td class="tdbg"><input name="t0" type="text" class="input" size="50" value="<%=t0%>"></td>
    </tr>
    <tr class="tdbg">
      <td align="right">�������ࣺ</td>
      <td><select name="t1"><option value="">��ѡ�����</option><%=Get_Class(t1)%></select></td>
    </tr>
	<tr class="tdbg">
      <td align="right">����ר�⣺</td>
      <td><select name="t33"><option value="0" <%=IIF(0=Clng(t33),"selected","")%>>������ר��</option>
	  <%Set Rs=Conn.Execute("Select ID,title From Sd_Topic Order By Id Desc"):While Not Rs.Eof%><option value="<%=Rs(0)%>" <%=IIF(Rs(0)=Clng(t33),"selected","")%>><%=Rs(1)%></option><%Rs.MoveNext:Wend%>
	  </select></td>
    </tr>
	<tr class="tdbg">
      <td align="right">�������ã�</td>
      <td><select name="t2"><option value="">��ѡ�����</option><option value="gb2312" <%=IIF(t2="gb2312","selected","")%>>Gb2312</option><option value="utf-8" <%=IIF(t2="utf-8","selected","")%>>Utf-8</option><option value="big5" <%=IIF(t2="big5","selected","")%>>Big5</option></select></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">Զ���б�URL��</td>
      <td class="tdbg"><input name="t3" type="text" class="input" size="50" value="<%=t3%>"></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�б��ҳ���ã�</td>
      <td class="tdbg"><input name="t4" type="radio" value="0" <%=IIF(t4=0,"checked","")%> onclick="<%For I=0 To 2%>$('#f<%=i%>')[0].style.display='none';<%Next%>" id="list01" /><label for="list01">��������</label> <input name="t4" type="radio" value="1" <%=IIF(t4=1,"checked","")%> onclick="<%For I=0 To 1%>$('#f<%=i%>')[0].style.display='block';<%Next%><%For I=2 To 2%>$('#f<%=i%>')[0].style.display='none';<%Next%>" id="list02" /><label for="list02">��������</label> <input name="t4" type="radio" value="2" <%=IIF(t4=2,"checked","")%> onclick="<%For I=2 To 2%>$('#f<%=i%>')[0].style.display='block';<%Next%><%For I=0 To 1%>$('#f<%=i%>')[0].style.display='none';<%Next%>" id="list03" /><label for="list03">�ֶ����</label> </td>
    </tr>
	<tr id="f0"<%IF t4<>1 Then%> class="dis"<%End IF%>>
      <td width="120" align="right" class="tdbg">�������ɣ�</td>
      <td class="tdbg"><input name="t5" type="text" class="input" size="50" value="<%=t5%>">��<span>��ʽ��http://www.website.com/list.asp?page={$ID}</span></td>
    </tr>
	<tr id="f1"<%IF t4<>1 Then%> class="dis"<%End IF%>>
      <td width="120" align="right" class="tdbg">���ɷ�Χ��</td>
      <td class="tdbg"><input name="t6" type="text" class="input" size="10" value="<%=t6%>"> �� <input name="t7" type="text" class="input" size="10" value="<%=t7%>">��<span>���磺1 - 9 ���� 9 - 1 ,ֻ�������֣���������߽���</span></td>
    </tr>
	<tr id="f2"<%IF t4<>2 Then%> class="dis"<%End IF%>>
      <td width="120" align="right" class="tdbg">�ֶ���ӣ�</td>
      <td class="tdbg"><textarea name="t8"  rows="6" class="inputs"><%=Content_Encode(t8)%></textarea><span><br>����һ����ַ�󰴻س�����������һ��</span></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�������ԣ�</td>
      <td class="tdbg"><input name="t11" type="checkbox" value="1" id="a01" <%=IIF(t11=1,"checked","")%> /><label for="a01">ͨ�����</label> <input name="t12" type="checkbox" value="1" id="a02" <%=IIF(t12=1,"checked","")%> /><label for="a02">����ͼƬ</label> <input name="t13" type="checkbox" value="1" id="a03" <%=IIF(t13=1,"checked","")%> <%IF Not Sdcms_Jpeg_t0 Then Echo "disabled" End IF%> /><label for="a03">ͼƬˮӡ</label> <input name="t14" type="checkbox" value="1" id="a04" <%=IIF(t14=1,"checked","")%> /><label for="a04">�����ɼ�</label> <input name="t15" type="checkbox" value="1" id="a05" <%=IIF(t15=1,"checked","")%> /><label for="a05">����ɼ�</label></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�߼����ԣ�</td>
      <td class="tdbg"><input name="t16" type="radio" value="0" id="top01" <%=IIF(t16=0,"checked","")%> onclick="<%For I=0 To 3%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="top01">����</label> <input name="t16" type="radio" value="1" onclick="<%For I=0 To 3%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="top02" <%=IIF(t16=1,"checked","")%> /><label for="top02">����</label></td>
    </tr>
	<tr id="t00"<%IF t16=0 Then%> class="dis"<%End IF%>>
      <td width="120" align="right" class="tdbg">��ǩ���ˣ�</td>
      <td class="tdbg"><input name="t17" type="checkbox" value="1" <%=IIF(t17=1,"checked","")%> id="t17_0" /><label for="t17_0">Iframe</label><input name="t18" type="checkbox" value="1" <%=IIF(t18=1,"checked","")%> id="t18_0" /><label for="t18_0">Object</label><input name="t19" type="checkbox" value="1" <%=IIF(t19=1,"checked","")%> id="t19_0" /><label for="t19_0">Script</label><input name="t20" type="checkbox" value="1" <%=IIF(t20=1,"checked","")%> id="t20_0" /><label for="t20_0">Div</label><input name="t21" type="checkbox" value="1" <%=IIF(t21=1,"checked","")%> id="t21_0" /><label for="t21_0">Class</label><input name="t22" type="checkbox" value="1" <%=IIF(t22=1,"checked","")%> id="t22_0" /><label for="t22_0">Table</label><input name="t23" type="checkbox" value="1" <%=IIF(t23=1,"checked","")%> id="t23_0" /><label for="t23_0">Tr</label><input name="t24" type="checkbox" value="1" <%=IIF(t24=1,"checked","")%> id="t24_0" /><label for="t24_0">Span</label><input name="t25" type="checkbox" value="1" <%=IIF(t25=1,"checked","")%> id="t25_0" /><label for="t25_0">Img</label><input name="t26" type="checkbox" value="1" <%=IIF(t26=1,"checked","")%> id="t26_0" /><label for="t26_0">Font</label><input name="t27" type="checkbox" value="1" <%=IIF(t27=1,"checked","")%> id="t27_0" /><label for="t27_0">A</label><input name="t28" type="checkbox" value="1" <%=IIF(t28=1,"checked","")%> id="t28_0" /><label for="t28_0">Html</label><input name="t29" type="checkbox" value="1" <%=IIF(t29=1,"checked","")%> id="t29_0" /><label for="t29_0">Td</label></td>
    </tr>
	<tr id="t01"<%IF t16=0 Then%> class="dis"<%End IF%>>
      <td width="120" align="right" class="tdbg">�����ַ��滻��</td>
      <td class="tdbg"><input type="button" class="bnt" name="addReplace" value="�����Ŀ" onclick="AddReplace();"> <input type="button" class="bnt" name="modifyReplace" value="�޸���Ŀ" onclick="return ModifyReplace();"> <input type="button" class="bnt" name="delReplace" value='ɾ����Ŀ' onclick="DelReplace();"> <input class="bnt" onClick="changepos(content,-1)"  value="�� ��" type="button"> <input class="bnt" onClick="changepos(content,1)" type="button" value="�� ��"><br><input type="hidden" name="t30" value="" >
	  <select class="inputs" name="content" style="width:500px;height:100px;margin-top:5px;" size="2" ondblclick="return ModifyReplace();" >
		<%
		IF Not IsNull(t30) Then
			Dim strReplaceArray
			strReplaceArray=Split(t30,",")
			For I = 0 To UBound(strReplaceArray)
				IF Len(strReplaceArray(i))>1 Then
					Echo"<option value="""&strReplaceArray(I)&""">"&Content_Encode(strReplaceArray(I))&"</option>"
				End IF
			Next	
		End If
		%>
	  </select><span style="position:absolute;">�����ִ�Сд</span></td>
    </tr>
	<tr id="t02"<%IF t16=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">�������ƣ�</td>
      <td class="tdbg"><input name="t31" type="text" class="input" size="50" value="<%=t31%>">��<span>0Ϊ�ɼ����гɹ�����</span></td>
    </tr>
	<tr id="t03"<%IF t16=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">�ˡ�������</td>
      <td class="tdbg"><input name="t32" type="text" class="input" size="50" value="<%=t32%>">��<span>�������ʼֵ��Ĭ��Ϊ0</span></td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��һ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Step2
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,Sql,Rs,i
Check_Add2
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
Sql="Select top 1 LsString,LoString,HsString,HoString,HttpUrlType,HttpUrlStr,x_tp,imhstr,imostr From "&Sd_Table&" Where Id="&ID
Set Rs=Coll_Conn.Execute(Sql)
IF Not Rs.Eof Then
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
t4=IsNum(t4,0)
t6=IsNum(t6,0)
Echo "��Ŀ���ã�<a href=""?action=edit&id="&id&""">��һ��</a>"
Echo " >> <a href=""?action=step2&id="&id&"""><b>�ڶ���</b></a> >> <a href=""?action=step3&id="&id&""">������</a>"
Echo "<br><br>"
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=step2_save&id=<%=id%>" onSubmit="return checkadd2()">
    <tr>
      <td align="right" class="tdbg"><b>Ч��Ԥ��</b>��</td>
      <td class="tdbg"><textarea rows="12" class="inputs"><%=Get_Url_Content(id)%></textarea></td>
    </tr>
    <tr>
      <td width="120" align="right" class="tdbg">�б�ʼ���룺</td>
      <td class="tdbg"><textarea name="t0" rows="4" class="inputs"><%=Content_Encode(t0)%></textarea></td>
    </tr>
   
	<tr>
      <td align="right" class="tdbg">�б�������룺</td>
      <td class="tdbg"><textarea name="t1" rows="4" class="inputs"><%=Content_Encode(t1)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���ӿ�ʼ���룺</td>
      <td class="tdbg"><textarea name="t2" rows="4" class="inputs"><%=Content_Encode(t2)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���ӽ������룺</td>
      <td class="tdbg"><textarea name="t3" rows="4" class="inputs"><%=Content_Encode(t3)%></textarea></td>
    </tr>
	<tr class="dis">
      <td align="right" class="tdbg">�ض����ӵ�ַ��</td>
      <td class="tdbg"><input name="t4" type="radio" value="0" id="top01" <%=IIF(t4=0,"checked","")%> onclick="$('#t00')[0].style.display='none';" /><label for="top01">�Զ�����</label> <input name="t4" type="radio" value="1" onclick="$('#t00')[0].style.display='block';" id="top02" <%=IIF(t4=1,"checked","")%> /><label for="top02">����λ��</label></td>
    </tr>
	<tr id="t00"<%IF t4=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">�ض������ַ���</td>
      <td class="tdbg"><input name="t5" type="text" class="input" size="50" value="<%=t5%>">��<span><br>��:javascript:Openwin("1001") 
function Openwin(ID) { popupWin = window.open('http://www.website.com/'+ ID + '.html','','width=400,height=300,scrollbars=yes')}
<br>��ȷ����:http://www.website.com/{$ID}.html </span></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�б�Сͼ���ã�</td>
      <td class="tdbg"><input name="t6" type="radio" value="0" id="top03" <%=IIF(t6=0,"checked","")%> onclick="<%For I=1 To 2%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="top03">��������</label> <input name="t6" type="radio" value="1" onclick="<%For I=1 To 2%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="top04" <%=IIF(t6=1,"checked","")%> /><label for="top04">ָ��λ��</label></td>
    </tr>
	<tr id="t01"<%IF t6=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">Сͼ��ʼ���룺</td>
      <td class="tdbg"><textarea name="t7" rows="4" class="inputs"><%=Content_Encode(t7)%></textarea></td>
    </tr>
	<tr id="t02"<%IF t6=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">Сͼ�������룺</td>
      <td class="tdbg"><textarea name="t8" rows="4" class="inputs"><%=Content_Encode(t8)%></textarea></td>
    </tr>
 
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��һ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Step3
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,Sql,Rs,I
Check_Add3
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
Sql="Select top 1 TsString,ToString,CsString,CoString,NewsPaingType,NPsString,NPoString,NewsUrlPaing_s,NewsUrlPaing_o,DateType,DsString,DoString,AuthorType"
Sql=Sql&",AsString,AoString,AuthorStr,CopyFromType,FsString,FoString,CopyFromStr,KeyType,KsString,KoString,KeyStr From "&Sd_Table&" Where Id="&ID
Set Rs=Coll_Conn.Execute(Sql)
IF Not Rs.Eof Then
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
	t22=Rs(22)
	t23=Rs(23)
End IF
t4=IsNum(t4,0)
t9=IsNum(t9,0)
t12=IsNum(t12,0)
t16=IsNum(t16,0)
t20=IsNum(t20,0)
Echo "��Ŀ���ã�<a href=""?action=edit&id="&id&""">��һ��</a>"
Echo " >> <a href=""?action=step2&id="&id&""">�ڶ���</a> >> <a href=""?action=step3&id="&id&"""><b>������</b></a>"
Echo "<br><br>"
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
  <form name="add" method="post" action="?action=step3_save&id=<%=id%>" onSubmit="return checkadd2()">
     <tr>
      <td align="right" class="tdbg"><b>Ŀ����ַ</b>��</td>
      <td class="tdbg"><div style="color:#00f;">���Ŀ����ַ�쳣��˵����һ�����������⣡</div><div style="height:200px;overflow-y:auto;border:1px dashed #ccc;padding:10px;"><%=Get_Urls(id)%></div></td>
    </tr>
	<tr>
      <td width="120" align="right" class="tdbg">���⿪ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t0" rows="4" class="inputs"><%=Content_Encode(t0)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���������ǣ�</td>
      <td class="tdbg"><textarea name="t1" rows="4" class="inputs"><%=Content_Encode(t1)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���Ŀ�ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t2" rows="4" class="inputs"><%=Content_Encode(t2)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���Ľ�����ǣ�</td>
      <td class="tdbg"><textarea name="t3" rows="4" class="inputs"><%=Content_Encode(t3)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">���ķ�ҳ���ã�</td>
      <td class="tdbg"><input name="t4" type="radio" value="0" id="top01" <%=IIF(t4=0,"checked","")%> onclick="<%For I=1 To 4%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="top01">��������</label> <input name="t4" type="radio" value="1" onclick="<%For I=1 To 4%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="top02" <%=IIF(t4=1,"checked","")%> /><label for="top02">ָ��λ��</label></td>
    </tr>
	<tr id="t01"<%IF t4=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��ҳ��ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t5" rows="4" class="inputs"><%=Content_Encode(t5)%></textarea></td>
    </tr>
	<tr id="t02"<%IF t4=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��ҳ������ǣ�</td>
      <td class="tdbg"><textarea name="t6" rows="4" class="inputs"><%=Content_Encode(t6)%></textarea></td>
    </tr>
    <tr id="t03"<%IF t4=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��ҳ���ӿ�ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t7" rows="4" class="inputs"><%=Content_Encode(t7)%></textarea></td>
    </tr>
	<tr id="t04"<%IF t4=0 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��ҳ���ӽ������:</td>
      <td class="tdbg"><textarea name="t8" rows="4" class="inputs"><%=Content_Encode(t8)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�������ã�</td>
      <td class="tdbg"><input name="t9" type="radio" value="0" id="t9_0" <%=IIF(t9=0,"checked","")%> onclick="<%For I=5 To 6%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="t9_0">��������</label> <input name="t9" type="radio" value="1" onclick="<%For I=5 To 6%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="t9_1" <%=IIF(t9=1,"checked","")%> /><label for="t9_1">���ñ�ǩ</label></td>
    </tr>
	<tr id="t05"<%IF t9<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">���ڿ�ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t10" rows="4" class="inputs"><%=Content_Encode(t10)%></textarea></td>
    </tr>
	<tr id="t06"<%IF t9<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">���ڽ�����ǣ�</td>
      <td class="tdbg"><textarea name="t11" rows="4" class="inputs"><%=Content_Encode(t11)%></textarea></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�������ã�</td>
      <td class="tdbg"><input name="t12" type="radio" value="0" id="t12_0" <%=IIF(t12=0,"checked","")%> onclick="<%For I=7 To 9%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="t12_0">��������</label> <input name="t12" type="radio" value="1" onclick="<%For I=7 To 8%>$('#t0<%=I%>')[0].style.display='block';<%Next%><%For I=9 To 9%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" id="t12_1" <%=IIF(t12=1,"checked","")%> /><label for="t12_1">���ñ�ǩ</label><input name="t12" type="radio" value="2" onclick="<%For I=7 To 8%>$('#t0<%=I%>')[0].style.display='none';<%Next%><%For I=9 To 9%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="t12_2" <%=IIF(t12=2,"checked","")%> /><label for="t12_2">ָ������</label></td>
    </tr>
	<tr id="t07"<%IF t12<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">���߿�ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t13" rows="4" class="inputs"><%=Content_Encode(t13)%></textarea></td>
    </tr>
	<tr id="t08"<%IF t12<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">���߽�����ǣ�</td>
      <td class="tdbg"><textarea name="t14" rows="4" class="inputs"><%=Content_Encode(t14)%></textarea></td>
    </tr>
	<tr id="t09"<%IF t12<>2 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">ָ�����ߣ�</td>
      <td class="tdbg"><input name="t15" type="text" class="input" size="50" value="<%=t15%>"></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">��Դ���ã�</td>
      <td class="tdbg"><input name="t16" type="radio" value="0" id="t16_0" <%=IIF(t16=0,"checked","")%> onclick="<%For I=10 To 12%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="t16_0">��������</label> <input name="t16" type="radio" value="1" onclick="<%For I=10 To 11%>$('#t0<%=I%>')[0].style.display='block';<%Next%><%For I=12 To 12%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" id="t16_1" <%=IIF(t16=1,"checked","")%> /><label for="t16_1">���ñ�ǩ</label><input name="t16" type="radio" value="2" onclick="<%For I=10 To 11%>$('#t0<%=I%>')[0].style.display='none';<%Next%><%For I=12 To 12%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="t16_2" <%=IIF(t16=2,"checked","")%> /><label for="t16_2">ָ����Դ</label></td>
    </tr>
	<tr id="t010"<%IF t16<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��Դ��ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t17" rows="4" class="inputs"><%=Content_Encode(t17)%></textarea></td>
    </tr>
	<tr id="t011"<%IF t16<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">��Դ������ǣ�</td>
      <td class="tdbg"><textarea name="t18" rows="4" class="inputs"><%=Content_Encode(t18)%></textarea></td>
    </tr>
	<tr id="t012"<%IF t16<>2 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">ָ����Դ��</td>
      <td class="tdbg"><input name="t19" type="text" class="input" size="50" value="<%=t19%>"></td>
    </tr>
	<tr>
      <td align="right" class="tdbg">�ؼ������ã�</td>
      <td class="tdbg"><input name="t20" type="radio" value="0" id="t20_0" <%=IIF(t20=0,"checked","")%> onclick="<%For I=13 To 15%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" /><label for="t20_0">��������</label> <input name="t20" type="radio" value="1" onclick="<%For I=13 To 14%>$('#t0<%=I%>')[0].style.display='block';<%Next%><%For I=15 To 15%>$('#t0<%=I%>')[0].style.display='none';<%Next%>" id="t20_1" <%=IIF(t20=1,"checked","")%> /><label for="t20_1">���ñ�ǩ</label><input name="t20" type="radio" value="2" onclick="<%For I=13 To 14%>$('#t0<%=I%>')[0].style.display='none';<%Next%><%For I=15 To 15%>$('#t0<%=I%>')[0].style.display='block';<%Next%>" id="t20_2" <%=IIF(t20=2,"checked","")%> /><label for="t20_2">ָ���ؼ���</label></td>
    </tr>
	<tr id="t013"<%IF t20<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">�ؼ��ֿ�ʼ��ǣ�</td>
      <td class="tdbg"><textarea name="t21" rows="4" class="inputs"><%=Content_Encode(t21)%></textarea></td>
    </tr>
	<tr id="t014"<%IF t20<>1 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">�ؼ��ֽ�����ǣ�</td>
      <td class="tdbg"><textarea name="t22" rows="4" class="inputs"><%=Content_Encode(t22)%></textarea></td>
    </tr>
	<tr id="t015"<%IF t20<>2 Then%> class="dis"<%End IF%>>
      <td align="right" class="tdbg">ָ���ؼ��֣�</td>
      <td class="tdbg"><input name="t23" type="text" class="input" size="50" value="<%=t23%>">��<span>�ؼ���֮����,�ָ����磺����,Sdcms </span></td>
    </tr>
    <tr class="tdbg">
	  <td>&nbsp;</td>
      <td><input type="submit" class="bnt" value="��һ��"> <input type="button" onClick="history.go(-1)" class="bnt" value="�� ��"></td>
    </tr>
	</form>
  </table>
<%
End Sub

Sub Step1_Save
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,t24,t25,t26,t27,t28,t29,t30,t31,t32,t33,Sql,Rs,LogMsg
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
t0=trim(request("t0"))
t1=trim(request("t1"))
t2=trim(request("t2"))
t3=trim(request("t3"))
t4=trim(request("t4"))
t5=trim(request("t5"))
t6=trim(request("t6"))
t7=trim(request("t7"))
t8=trim(request("t8"))
t9=trim(request("t9"))
t10=trim(request("t10"))
t11=trim(request("t11"))
t12=trim(request("t12"))
t13=trim(request("t13"))
t14=trim(request("t14"))
t15=trim(request("t15"))
t16=trim(request("t16"))
t17=trim(request("t17"))
t18=trim(request("t18"))
t19=trim(request("t19"))
t20=trim(request("t20"))
t21=trim(request("t21"))
t22=trim(request("t22"))
t23=trim(request("t23"))
t24=trim(request("t24"))
t25=trim(request("t25"))
t26=trim(request("t26"))
t27=trim(request("t27"))
t28=trim(request("t28"))
t29=trim(request("t29"))
t30=trim(request("t30"))
t31=trim(request("t31"))
t32=trim(request("t32"))
t33=trim(request("t33"))

t4=IsNum(t4,0)
t16=IsNum(t16,0)
t31=IsNum(t31,0)
t32=IsNum(t32,0)
t33=IsNum(t33,0)

IF t4=1 Then
	IF Len(t5)=0 Then
		Alert "�������ɲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t6)=0 Then
		Alert "���ɷ�Χ����Ϊ�գ��ұ���Ϊ����","javascript:history.go(-1)":Died
	End IF
	IF Len(t7)=0 Then
		Alert "���ɷ�Χ����Ϊ�գ��ұ���Ϊ����","javascript:history.go(-1)":Died
	End IF
End IF

IF t4=2 Then
	IF Len(t8)=0 Then
		Alert "�ֶ���Ӳ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t4=3 Then
	IF Len(t9)=0 Then
		Alert "��ҳ��ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t10)=0 Then
		Alert "��ҳ������ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF
t8=Check_event(Re(t8,Chr(13),"|"),"|","")
t30=Check_event(t30,",","")
IF Len(t30)>0 Then
	IF Instr(t30,"|")=0 Then
		Alert "�����ַ��滻��ʽ����","javascript:history.go(-1)":Died
	End IF
End IF

IF ID=0 Then sdcms.Check_lever 21 Else sdcms.Check_lever 22
Set rs=Server.CreateObject("adodb.recordset")
Sql="Select top 1 ItemName,Classid,selEncoding,ListStr,ListPaingType,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,LPsString,LPoString"
Sql=Sql&",Passed,SaveFiles,Thumb_WaterMark,Flag,CollecOrder,Coll_Top,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class"
Sql=Sql&",Script_table,Script_tr,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,Script_Td,strReplace,CollecNewsNum,Hits,SpecialID,ID From "&Sd_Table&" "
IF Id>0 then 
	Sql=Sql&" where ID="&ID&""
End IF
Rs.Open Sql,Coll_Conn,1,3

IF Id=0 then 
  Rs.Addnew
Else
	IF Rs.Eof Then
		Alert "��������","javascript:history.go(-1)":Died
	End IF
	Rs.Update
End IF
	Rs(0)=Left(t0,50)
	Rs(1)=IsNum(t1,0)
	Rs(2)=Left(t2,50)
	Rs(3)=t3
	Rs(4)=IsNum(t4,0)
	Rs(5)=Left(t5,150)
	Rs(6)=IsNum(t6,0)
	Rs(7)=IsNum(t7,0)
	Rs(8)=t8
	'Rs(9)=t9
	'Rs(10)=t10
	Rs(11)=IsNum(t11,0)
	Rs(12)=IsNum(t12,0)
	Rs(13)=IsNum(t13,0)
	Rs(14)=IsNum(t14,0)
	Rs(15)=IsNum(t15,0)
	Rs(16)=IsNum(t16,0)
	Rs(17)=IsNum(t17,0)
	Rs(18)=IsNum(t18,0)
	Rs(19)=IsNum(t19,0)
	Rs(20)=IsNum(t20,0)
	Rs(21)=IsNum(t21,0)
	Rs(22)=IsNum(t22,0)
	Rs(23)=IsNum(t23,0)
	Rs(24)=IsNum(t24,0)
	Rs(25)=IsNum(t25,0)
	Rs(26)=IsNum(t26,0)
	Rs(27)=IsNum(t27,0)
	Rs(28)=IsNum(t28,0)
	Rs(29)=IsNum(t29,0)
	Rs(30)=t30
	Rs(31)=IsNum(t31,0)
	Rs(32)=IsNum(t32,0)
	Rs(33)=IsNum(t33,0)
	ID=Rs(34)
Rs.Update
Del_Cache("Get_Coll_List_"&ID):Del_Cache("Coll_Pic_List_"&ID):Del_Cache("Get_Info_Config_"&ID)
IF ID=0 Then LogMsg="�����Ŀ" Else LogMsg="�޸���Ŀ"
AddLog sdcms_adminname,GetIp,LogMsg&t0,0
Go("?Action=Step2&ID="&ID&"")
End Sub

Sub Step2_Save
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,rs,sql
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
t0=trim(request("t0"))
t1=trim(request("t1"))
t2=trim(request("t2"))
t3=trim(request("t3"))
t4=trim(request("t4"))
t5=trim(request("t5"))
t6=trim(request("t6"))
t7=trim(request("t7"))
t8=trim(request("t8"))
t4=IsNum(t4,0)
t6=IsNum(t6,0)

IF t4>0 Then
	IF Len(t5)=0 Then
		Alert "�ض������ַ�����Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t6>0 Then
	IF Len(t7)=0 Then
		Alert "Сͼ��ʼ���벻��Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t8)=0 Then
		Alert "Сͼ�������벻��Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

Set Rs=Server.CreateObject("adodb.recordset")
Sql="Select top 1 LsString,LoString,HsString,HoString,HttpUrlType,HttpUrlStr,x_tp,imhstr,imostr From "&Sd_Table&" Where Id="&ID
Rs.Open Sql,Coll_Conn,1,3
IF Rs.Eof Then
	Alert "��������","javascript:history.go(-1)":Died
End IF
Rs.Update
	Rs(0)=t0
	Rs(1)=t1
	Rs(2)=t2
	Rs(3)=t3
	Rs(4)=t4
	Rs(5)=Left(t5,255)
	Rs(6)=t6
	Rs(7)=t7
	Rs(8)=t8
Rs.Update
Del_Cache("Get_Coll_List_"&ID):Del_Cache("Coll_Pic_List_"&ID):Del_Cache("Get_Info_Config_"&ID)
Go("?Action=Step3&ID="&ID&"")
End Sub

Sub Step3_Save
Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,rs,sql
Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
t0=trim(request("t0"))
t1=trim(request("t1"))
t2=trim(request("t2"))
t3=trim(request("t3"))
t4=trim(request("t4"))
t5=trim(request("t5"))
t6=trim(request("t6"))
t7=trim(request("t7"))
t8=trim(request("t8"))
t9=trim(request("t9"))
t10=trim(request("t10"))
t11=trim(request("t11"))
t12=trim(request("t12"))
t13=trim(request("t13"))
t14=trim(request("t14"))
t15=trim(request("t15"))
t16=trim(request("t16"))
t17=trim(request("t17"))
t18=trim(request("t18"))
t19=trim(request("t19"))
t20=trim(request("t20"))
t21=trim(request("t21"))
t22=trim(request("t22"))
t23=trim(request("t23"))
t4=IsNum(t4,0)
t9=IsNum(t9,0)
t12=IsNum(t12,0)
t16=IsNum(t16,0)
t20=IsNum(t20,0)

IF t4>0 Then
	IF Len(t5)=0 Then
		Alert "��ҳ��ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t6)=0 Then
		Alert "��ҳ������ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t7)=0 Then
		Alert "��ҳ���ӿ�ʼ���","javascript:history.go(-1)":Died
	End IF
	IF Len(t8)=0 Then
		Alert "��ҳ���ӽ������","javascript:history.go(-1)":Died
	End IF
End IF

IF t9>0 Then
	IF Len(t10)=0 Then
		Alert "���ڿ�ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t11)=0 Then
		Alert "���ڽ�����ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t12=1 Then
	IF Len(t13)=0 Then
		Alert "���߿�ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t14)=0 Then
		Alert "���߽�����ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t12=2 Then
	IF Len(t15)=0 Then
		Alert "ָ�����߲���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t16=1 Then
	IF Len(t17)=0 Then
		Alert "��Դ��ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t18)=0 Then
		Alert "��Դ������ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t16=2 Then
	IF Len(t19)=0 Then
		Alert "ָ����Դ����Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t20=1 Then
	IF Len(t21)=0 Then
		Alert "�ؼ��ֿ�ʼ��ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
	IF Len(t22)=0 Then
		Alert "�ؼ��ֽ�����ǲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

IF t20=2 Then
	IF Len(t23)=0 Then
		Alert "ָ���ؼ��ֲ���Ϊ��","javascript:history.go(-1)":Died
	End IF
End IF

Set Rs=Server.CreateObject("adodb.recordset")
Sql="Select top 1 TsString,ToString,CsString,CoString,NewsPaingType,NPsString,NPoString,NewsUrlPaing_s,NewsUrlPaing_o,DateType,DsString,DoString,AuthorType"
Sql=Sql&",AsString,AoString,AuthorStr,CopyFromType,FsString,FoString,CopyFromStr,KeyType,KsString,KoString,KeyStr From "&Sd_Table&" Where Id="&ID
Rs.Open Sql,Coll_Conn,1,3
IF Rs.Eof Then
	Alert "��������","javascript:history.go(-1)":Died
End IF
	Rs.Update
	Rs(0)=t0
	Rs(1)=t1
	Rs(2)=t2
	Rs(3)=t3
	Rs(4)=IsNum(t4,0)
	Rs(5)=t5
	Rs(6)=t6
	Rs(7)=t7
	Rs(8)=t8
	Rs(9)=IsNum(t9,0)
	Rs(10)=t10
	Rs(11)=t11
	Rs(12)=IsNum(t12,0)
	Rs(13)=t13
	Rs(14)=t14
	Rs(15)=t15
	Rs(16)=IsNum(t16,0)
	Rs(17)=t17
	Rs(18)=t18
	Rs(19)=t19
	Rs(20)=IsNum(t20,0)
	Rs(21)=t21
	Rs(22)=t22
	Rs(23)=t23
	Rs.Update
	Del_Cache("Get_Coll_List_"&ID):Del_Cache("Coll_Pic_List_"&ID):Del_Cache("Get_Info_Config_"&ID)
	Alert "�������\n\n����ɼ�ǰ�Ȳ����¿����Ƿ�������","?"
End Sub

Sub Copy
	Dim ID:ID=IsNum(Trim(Request.QueryString("ID")),0)
	IF ID="" Then
		  Alert "��ѡ��Ҫ���Ƶ���Ŀ��","?":Died
	Else
		Dim copy_s,copyitem(999),Sql,Rs,i
		Sql="Select top 1 * from "&Sd_Table&" where ID="&ID
		Set Rs=Server.CreateObject("adodb.recordset")
		Rs.Open Sql,Coll_Conn,3,3
		copy_s=Cint(Rs.Fields.count)
		For I=1 To copy_s-1
		copyitem(I)=Rs(I)
		Next
		Rs.AddNew
		For I=1 To copy_s-1
		Rs(I)=copyitem(I)
		Rs(1)=copyitem(1)&"_����"
		Next
		Rs.Update
		Rs.Close
		Go("?")
	End If
End Sub

Sub Del
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	IF Len(ID)>0 Then
		AddLog sdcms_adminname,GetIp,"ɾ���ɼ���Ŀ�����Ϊ"&ID,0
		Coll_Conn.Execute("Delete From "&Sd_Table&" where id in("&ID&")")
		Coll_Conn.Execute("Delete From Sd_Coll_Filters where ItemID in("&ID&")")
		Coll_Conn.Execute("Delete From Sd_Coll_History where ItemID in("&ID&")")
	End IF
	Go("?") 
End Sub

Sub Pass(t0)
	Dim ID:ID=Trim(Request("ID"))
	ID=Re(ID," ","")
	Coll_Conn.Execute("Update "&Sd_Table&" Set flag="&t0&" where id in("&ID&")")
	Del_Cache("Get_Coll_List_"&ID):Del_Cache("Coll_Pic_List_"&ID):Del_Cache("Get_Info_Config_"&ID)
	Go("?") 
End Sub

Sub Import
Check_Import
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
<form action="?action=importnext" method="post" onsubmit="return checkimport(this)">
    <tr>
        <td width="120" align="center" class="tdbg">����Դ�� </td>
        <td class="tdbg"><input name="t0" type="text" class="input" value="../<%=Sdcms_DataFile%>/Coll_Item.mdb" size="30" /></td>
    </tr>
    <tr class="tdbg">
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" class="bnt" value="��һ��" /></td>
    </tr>
</form>
</table>
<%
End Sub

Sub ImportNext
	Dim t0
	t0=FilterHtml(Trim(Request.Form("t0")))
	IF Len(t0)=0 Then
		Alert "����ѡ������Դ","javascript:history.go(-1)":Died
	End IF
	On Error ReSume Next
	Dim Conn2
	Set Conn2=Server.CreateObject("ADODB.Connection")
	Conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(t0)
	IF Err Then
		Alert "����Դ����������ѡ��","javascript:history.go(-1)":Died
		Err.Clear
	End IF
	Dim Rs
	Check_Import_Next
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
<form action="?action=importover&t0=<%=t0%>" method="post" onsubmit="return checkimportnext(this)">
    <tr>
        <td width="120" align="center" class="tdbg">ѡ����Ŀ��<br><input  type="checkbox" name="checkall" id="checkall" onclick="checkselect(this,$('#t1')[0])"><label for="checkall">ȫѡ/ȡ��</label> </td>
        <td class="tdbg"><select name="t1" id="t1" size="1" multiple="multiple" style="width:400px;height:250px;">
			<%
            Set Rs=Conn2.Execute("Select ID,ItemName From "&Sd_Table&" Order By ID Desc")
			IF Rs.Eof Then
			%>
            <option value="">û�пɵ������Ŀ</option>
            <%
			End IF
            While Not Rs.Eof
            %>
            <option value="<%=Rs(0)%>"><%=Rs(1)%></option>
            <%Rs.MoveNext:Wend:Rs.Close:Set Rs=Nothing%>
        </select></td>
    </tr>
    <tr class="tdbg">
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" class="bnt" value="��������" /></td>
    </tr>
</form>
</table>
<%
End Sub

Sub ImportOver
	Dim t0,t1
	t0=FilterHtml(Trim(Request.QueryString("t0")))
	t1=Re(FilterHtml(Trim(Request.Form("t1")))," ","")
	IF Len(t0)=0 Then
		Alert "����ѡ������Դ","javascript:history.go(-1)":Died
	End IF
	IF Len(t1)=0 Then
		Alert "����ѡ��Ҫ�������Ŀ","javascript:history.go(-1)":Died
	End IF
	On Error ReSume Next
	Dim Conn2
	Set Conn2=Server.CreateObject("ADODB.Connection")
	Conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(t0)
	IF Err Then
		Alert "����Դ����������ѡ��","javascript:history.go(-1)":Died
		Err.Clear
	End IF
	Dim Rs,Rs1
	
	Dim copy_s,copyitem(999),Sql,I
	Sql="Select * From "&Sd_Table&" where ID In("&t1&")"
	Set Rs1=Server.CreateObject("Adodb.Recordset")
	Rs1.Open Sql,Conn2,1,3
	While Not Rs1.Eof
		copy_s=Cint(Rs1.Fields.Count)
		For I=1 To copy_s-1
			copyitem(I)=Rs1(I)
			
		Next
		Set Rs=Server.CreateObject("Adodb.Recordset")
		Rs.Open "Select * From Sd_Coll_Item",Coll_Conn,1,3
		Rs.AddNew
			For I=1 To copy_s-1
				Rs(I)=copyitem(I)
			Next
		Rs.Update
		Rs.Close
	Rs1.MoveNext
	Wend
	Alert "���ݵ���ɹ�","?"
End Sub

Sub Export
Check_Import
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
<form action="?action=Exportnext" method="post" onsubmit="return checkexport(this)">
    <tr>
        <td width="120" align="center" class="tdbg">ѡ����Ŀ��<br><input  type="checkbox" name="checkall" id="checkall" onclick="checkselect(this,$('#t0')[0])"><label for="checkall">ȫѡ/ȡ��</label> </td>
        <td class="tdbg"><select name="t0" id="t0" size="1" multiple="multiple" style="width:400px;height:250px;">
			<%
            Set Rs=Coll_Conn.Execute("Select ID,ItemName From "&Sd_Table&" Order By ID Desc")
			IF Rs.Eof Then
			%>
            <option value="">û�пɵ�������Ŀ</option>
            <%
			End IF
            While Not Rs.Eof
            %>
            <option value="<%=Rs(0)%>"><%=Rs(1)%></option>
            <%Rs.MoveNext:Wend:Rs.Close:Set Rs=Nothing%>
        </select></td>
    </tr>
    <tr class="tdbg">
        <td align="center">Ŀ�����ݿ�</td>
        <td><input name="t1" type="text" class="input" value="../<%=Sdcms_DataFile%>/Coll_Item.mdb" size="30" /><input type="checkbox" value="1" name="t2" id="t2" checked="checked" /><label for="t2">���Ŀ�����ݿ�</label></td>
    </tr>
    <tr class="tdbg">
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" class="bnt" value="��������" /></td>
    </tr>
</form>
</table>
<%
End Sub

Sub Exportnext
	Dim t0,t1,t2
	t0=FilterHtml(Trim(Request.Form("t0")))
	t1=Re(FilterHtml(Trim(Request.Form("t1")))," ","")
	t2=IsNum(Trim(Request.Form("t2")),0)
	IF Len(t0)=0 Then
		Alert "����ѡ��Ҫ��������Ŀ","javascript:history.go(-1)":Died
	End IF
	IF Len(t1)=0 Then
		Alert "����ѡ��Ŀ������Դ","javascript:history.go(-1)":Died
	End IF
	On Error ReSume Next
	Dim Conn2
	Set Conn2=Server.CreateObject("ADODB.Connection")
	Conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(t1)
	IF Err Then
		Alert "Ŀ������Դ����������ѡ��","javascript:history.go(-1)":Died
		Err.Clear
	End IF
	Dim Rs,Rs1
	IF t2=1 Then
		Conn2.Execute("Delete From "&Sd_Table&" ")
	End IF
	Dim copy_s,copyitem(999),Sql,I
	Sql="Select * From "&Sd_Table&" where ID In("&t0&")"
	Set Rs1=Server.CreateObject("Adodb.Recordset")
	Rs1.Open Sql,Coll_Conn,1,3
	While Not Rs1.Eof
		copy_s=Cint(Rs1.Fields.Count)
		For I=1 To copy_s-1
			copyitem(I)=Rs1(I)
		Next
		Set Rs=Server.CreateObject("Adodb.Recordset")
		Rs.Open "Select * From Sd_Coll_Item",Conn2,1,3
		Rs.AddNew
			For I=1 To copy_s-1
				Rs(I)=copyitem(I)
			Next
		Rs.Update
		Rs.Close
	Rs1.MoveNext
	Wend
	Alert "���ݵ����ɹ�","?"
End Sub

Sub Check_Info
Echo("<script>")
Echo("function changepos(obj,index)")
Echo("{")
Echo("if(index==-1){")
Echo("if (obj.selectedIndex>0){")
Echo("obj.options(obj.selectedIndex).swapNode(obj.options(obj.selectedIndex-1))")
Echo("}")
Echo("}")
Echo("else if(index==1){")
Echo("if (obj.selectedIndex<obj.options.length-1){")
Echo("obj.options(obj.selectedIndex).swapNode(obj.options(obj.selectedIndex+1))")
Echo("}")
Echo("}")
Echo("}")

Echo("function AddReplace(){")
Echo("  var thisReplace='�滻ǰ���ַ���'+(document.add.content.length+1)+'|�滻����ַ���'+(document.add.content.length+1); ")
Echo("  var Replace=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',thisReplace);")
Echo("  if(Replace!=null&&Replace!=''){document.add.content.options[document.add.content.length]=new Option(Replace,Replace);}")
Echo("}")
Echo("function ModifyReplace(){")
Echo("  if(document.add.content.length==0) return false;")
Echo("  var thisReplace=document.add.content.value; ")
Echo("  if (thisReplace=='') {alert('����ѡ��һ����Ŀ���ٵ��޸İ�ť��');return false;}")
Echo("  var Replace=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',thisReplace);")
Echo("  if(Replace!=thisReplace&&Replace!=null&&Replace!=''){document.add.content.options[document.add.content.selectedIndex]=new Option(Replace,Replace);}")
Echo("}")
Echo("function DelReplace(){")
Echo("  if(document.add.content.length==0) return false;")
Echo("  var thisReplace=document.add.content.value; ")
Echo("  if (thisReplace=='') {alert('����ѡ��һ����Ŀ���ٵ�ɾ����ť��');return false;}")
Echo("  document.add.content.options[document.add.content.selectedIndex]=null;")
Echo("}")
Echo("	function checkadd()")
Echo("	{")
Echo("	if (document.add.t0.value=='')")
Echo("	{")
Echo("	alert('��Ŀ���Ʋ���Ϊ��');")
Echo("	document.add.t0.focus();")
Echo("	return false")
Echo("	}")
Echo("	if (document.add.t1.value=='')")
Echo("	{")
Echo("	alert('�������಻��Ϊ��');")
Echo("	document.add.t1.focus();")
Echo("	return false")
Echo("	}")
Echo("	if (document.add.t2.value=='')")
Echo("	{")
Echo("	alert('�������ò���Ϊ��');")
Echo("	document.add.t2.focus();")
Echo("	return false")
Echo("	}")
Echo("	if (document.add.t3.value=='')")
Echo("	{")
Echo("	alert('Զ���б�URL����Ϊ��');")
Echo("	document.add.t3.focus();")
Echo("	return false")
Echo("	}") 
Echo("var s=""""; ")
Echo("for(i=0;i<=document.all(""content"").length-1;i++) ")
Echo("{ ")
Echo("var s=s+document.add.content.options(i).value+"",""; ")
Echo("} ")
Echo("document.add.t30.value=s")
Echo("	}")
Echo("	</script>")
End Sub

Sub Check_Add2
	Echo("<script>")
	Echo("	function checkadd2()")
	Echo("	{")
	Echo("	if (document.add.t0.value=='')")
	Echo("	{")
	Echo("	alert('�б�ʼ���벻��Ϊ��');")
	Echo("	document.add.t0.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t1.value=='')")
	Echo("	{")
	Echo("	alert('�б�������벻��Ϊ��');")
	Echo("	document.add.t1.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t2.value=='')")
	Echo("	{")
	Echo("	alert('���ӿ�ʼ���벻��Ϊ��');")
	Echo("	document.add.t2.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t3.value=='')")
	Echo("	{")
	Echo("	alert('���ӽ������벻��Ϊ��');")
	Echo("	document.add.t3.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	}")
	Echo("	</script>")
End Sub

Sub Check_Add3
	Echo("<script>")
	Echo("	function checkadd2()")
	Echo("	{")
	Echo("	if (document.add.t0.value=='')")
	Echo("	{")
	Echo("	alert('���⿪ʼ��ǲ���Ϊ��');")
	Echo("	document.add.t0.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t1.value=='')")
	Echo("	{")
	Echo("	alert('���������ǲ���Ϊ��');")
	Echo("	document.add.t1.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t2.value=='')")
	Echo("	{")
	Echo("	alert('���ӿ�ʼ��ǲ���Ϊ��');")
	Echo("	document.add.t2.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	if (document.add.t3.value=='')")
	Echo("	{")
	Echo("	alert('���ӽ�����ǲ���Ϊ��');")
	Echo("	document.add.t3.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	}")
	Echo("	</script>")
End Sub

Sub Check_Import
	Echo("<script>")
	Echo("	function checkimport(theform)")
	Echo("	{")
	Echo("	if (theform.t0.value=='')")
	Echo("	{")
	Echo("	alert('����������Դ');")
	Echo("	theform.t0.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	}")
	Echo("	</script>")
End Sub

Sub Check_Import_Next
	Echo("<script>")
	Echo("	function checkimportnext(theform)")
	Echo("	{")
	Echo("	if (theform.t1.value=='')")
	Echo("	{")
	Echo("	alert('��ѡ��Ҫ�������Ŀ');")
	Echo("	theform.t1.focus();")
	Echo("	return false")
	Echo("	}")
	Echo("	}")
	Echo("	</script>")
End Sub


Sub Check_Import
	Echo("<script>")
	Echo("	function checkexport(theform)")
	Echo("	{")
		Echo("	if (theform.t0.value=='')")
		Echo("	{")
		Echo("	alert('��ѡ��Ҫ��������Ŀ');")
		Echo("	theform.t0.focus();")
		Echo("	return false")
		Echo("	}")
		Echo("	if (theform.t1.value=='')")
		Echo("	{")
		Echo("	alert('������Ŀ������Դ');")
		Echo("	theform.t1.focus();")
		Echo("	return false")
		Echo("	}")
	Echo("	}")
	Echo("	</script>")
End Sub
%>  
</div>
</body>
</html>