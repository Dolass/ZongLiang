<!--#include file="sdcms_check.asp"-->
<%
Dim sdcms,sdcms_c,stype
Set sdcms=New Sdcms_Admin
sdcms.Check_admin
sdcms.Check_lever 31
Sdcms_Head
Stype=Trim(Request("Stype"))
IF Sdcms_Mode<>2 Then
	NoHtml
Else
	Select Case Stype
		Case "1":Make_Html_1:Create_Msg
		Case "2":Make_Html_2:Create_Msg
		Case "3":Make_Html_3:Create_Msg
		Case "4":Make_Html_4:Create_Msg
		Case "5":Make_Html_5:Create_Msg
		Case Else:Make_html
	End Select
	CloseDb
End IF

Sub Make_html
	Dim Act:Act=Trim(Request.QueryString("Act"))

	Set Sdcms_C=New Sdcms_Create
		Select Case Act
			Case "1":Make_Index
			Case "2":Make_Class
			Case "3":Make_Info
			Case "4":Make_Page
			Case "5":Make_Map
			Case "go":Make_Class_Arr
			Case "pagelist":Make_Class_Page
			Case Else:Echo "��������":Died
		End Select
		Make_Time
	Set Sdcms_C=Nothing
End Sub

Sub Make_Index
	Sdcms_c.Create_index()
End Sub

Sub Make_Class
	Dim t0,t1
	t0=Trim(Request.QueryString("t0"))
	IF Len(t0)=0 Then Echo "����ѡ��һ������<br>":Died
	t0=Replace(t0," ","")
	t1=Split(t0,",")
	
	Dim Total_Class_Num
	Total_Class_Num=Ubound(t1)+1
	Add_Cookies "ClassIDArray",t0
	Go "?act=go&Total_Class_Num="&Total_Class_Num&"&This_Arr=1"
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
			Go "?act=pagelist&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"&id="&This_ID&"&TotalPage="&TotalPage&"&page=1":Died
		End IF
	End IF
	Rs.Close
	Set Rs=Nothing
	'======================================
	This_Arr=This_Arr+1
	
	IF This_Arr<=Total_Class_Num Then
		Echo "<script>setTimeout(""location.href='?act=go&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"';"",""200"");</script>"
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
		Echo "<script>setTimeout(""location.href='?act=pagelist&id="&id&"&TotalPage="&TotalPage&"&page="&Pages&"&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr&"';"",""100"");</script>"
	Else
		IF This_Arr>=Total_Class_Num Then
			Echo "<br><b>�������</b><br>":Exit Sub
		End IF
		Echo "<script>setTimeout(""location.href='?act=go&Total_Class_Num="&Total_Class_Num&"&This_Arr="&This_Arr+1&"';"",""100"");</script>"
	End IF

End Sub

Sub Make_Info
	Dim t0,t1,t2,t3,t4,t5,Rs,Lastid
	t0=Clng(Trim(Request.QueryString("t0")))
	t1=IsNum(Trim(Request.QueryString("t1")),0)
	t2=Trim(Request.QueryString("t2"))
	t3=IsNum(Trim(Request.QueryString("t3")),0)
	IF t0>0 Then t5=" classid="&t0&" and "
    IF t1>0 Then t4=" top "&t1
	Set Rs=Conn.Execute("select "&t4&" id from sd_info where "&t5&" id>"&t3&" and Ispass=1 order by id")
	IF Rs.Eof Then
		Echo "û����Ϣ��������<br>"
		Lastid=0
	End IF
	While Not Rs.Eof
		sdcms_c.Create_info_show Rs(0)
		echo "<script>window.scroll(0,999999)</script>"
		Lastid=rs(0)
		Response.Flush()
	Rs.Movenext
	Wend
	Rs.Close
	Set Rs=Nothing
	Set Rs=Conn.Execute("select top 1 id from sd_info where "&t5&" id>"&Lastid&" and Ispass=1 order by id")
	IF Not Rs.Eof Then
		Echo "<div id=""Info_Load"">"&t2&"�����������������Ϣ</div>" & VbCrLf
		Response.Flush()
		Echo "<script language=JavaScript>"
		Echo "var secs="&t2&";var wait=secs * 1000;"
		Echo "for(i=1; i<=secs;i++){window.setTimeout(""Update("" + i + "")"", i * 1000);}"
		Echo "function Update(num){if(num != secs){printnr = (wait / 1000) - num;"
		Echo "$(""#Info_Load"")[0].style.width=(num/secs)*100+""%"";"
		Echo "$(""#Info_Load"").html(""ʣ��""+printnr+""��"");}}"
		Echo "setTimeout(""window.location='?Act=3&t0="&t0&"&t1="&t1&"&t2="&t2&"&t3="&Lastid&"'"","&Int(t2*1000)&");</script>"
		Response.Flush()
	Else
		Echo "�������<br>"
		echo "<script>window.scroll(0,999999)</script>"
	End IF
	Rs.Close
	Set Rs=Nothing
End Sub

Sub Make_Page
	Dim Rs
	Set Rs=Conn.Execute("select id from sd_other order by id desc")
	IF Rs.Eof Then
		Echo "û�е�ҳ��Ҫ����<br>"
	End IF
	While Not Rs.Eof
		sdcms_c.Create_other Rs(0)
		echo "<script>window.scroll(0,999999)</script>"
		Response.Flush()
	Rs.MoveNext
	Wend
	Rs.Close
	Set Rs=Nothing
End Sub

Sub Make_Map
	Dim t0,t1,t2,t3,t4
	t0=Trim(Request.QueryString("t0"))
	t1=Trim(Request.QueryString("t1"))
	t2=Trim(Request.QueryString("t2"))
	t3=Trim(Request.QueryString("t3"))
	t4=Trim(Request.QueryString("t4"))
	IF Len(t0)>0 Then
	sdcms_c.Create_map
	End IF
	IF Len(t1)>0 Then
	sdcms_c.Create_google_map t2,t3,t4
	End IF
	IF Len(t0&t1)=0 Then Echo "û��ѡ���κ���Ŀ<br>"
End Sub

Sub Make_Time
	Echo "<br>��ʱ��"&Runtime&" ��"
	echo "<script>window.scroll(0,999999)</script>"
End Sub

Sub Make_Html_1
%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
  <form action="?" target="Loadinfo">
    <tr>
      <td class="tdbg01">������ҳ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><input type="hidden" name="act" value="1" /><input name="Submit" type="submit" class="bnt" value="��ʼ����" /></td>
    </tr>
	</form>
  </table>
<%End Sub:Sub Make_Html_2%> 
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <form action="?" target="Loadinfo">
	<tr>
      <td colspan="2" class="tdbg01">������Ŀ</td>
    </tr>
    <tr>
      <td width="120" height="25" align="center" class="tdbg">��Ŀѡ��<br /><input  type="checkbox" name="checkall" id="checkall"  onclick="checkselect(this,$('#t0')[0])"><label for="checkall">ȫѡ/ȡ��</label> 
</td>
      <td class="tdbg"><select name="t0" size="10" multiple="multiple" id="t0" style="width:60%;"><%=Get_Class(0)%></select><br />
<span>֧��ʹ�� Ctrl �� Shift ����ѡ</span></td>
    </tr>
    <tr>
      <td height="25" colspan="2" class="tdbg"><input type="hidden" name="act" value="2" /><input name="Submit" type="submit" class="bnt" value="��ʼ����" /></td>
    </tr>
	</form>
  </table>
<%End Sub:Sub Make_Html_3%>  
   <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <form action="?" target="Loadinfo">
	<tr>
      <td colspan="2" class="tdbg01">������Ϣ</td>
    </tr>
    <tr>
      <td width="120" height="25" align="center" class="tdbg">��Ŀѡ��</td>
      <td class="tdbg"><select name="t0" size="1" id="t0"><option value="0">������Ŀ</option><%=Get_Class(0)%></select></td>
    </tr>
	<tr>
      <td height="25" align="center" class="tdbg">ÿ��������</td>
      <td class="tdbg"><input name="t1" type="text" value="20" class="input" />��<span>Ϊ0ʱ����ȫ��</span></td>
    </tr>
	<tr>
      <td height="25" align="center" class="tdbg">ÿ�μ����</td>
      <td class="tdbg"><input name="t2" type="text" value="5" class="input" />��<span>��λ����</span></td>
    </tr>
    <tr>
      <td height="25" colspan="2" class="tdbg"><input type="hidden" name="act" value="3" /><input name="Submit" type="submit" class="bnt" value="��ʼ����" /></td>
    </tr>
	</form>
  </table>
<%End Sub:Sub Make_Html_4%>  
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <form action="?" target="Loadinfo">
	<tr>
      <td class="tdbg01">���ɵ�ҳ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg"><input type="hidden" name="act" value="4" /><input name="Submit" type="submit" class="bnt" value="��ʼ����" /></td>
    </tr>
	</form>
  </table>
<%End Sub:Sub Make_Html_5%>
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <form action="?" target="Loadinfo">
	<tr>
      <td colspan="2" class="tdbg01">���ɵ�ͼ</td>
    </tr>
    <tr>
      <td width="120" height="25" align="center" class="tdbg">��ͼѡ�</td>
      <td class="tdbg"><input name="t0" type="checkbox" value="1" checked="checked" id="t0" /><label for="t0">HTML��ͼ</label> <input name="t1" type="checkbox" value="1" checked="checked"  id="t1" /><label for="t1">Google��ͼ</label>    
      </td>
    </tr>
	<tr>
      <td height="25" align="center" class="tdbg">Google������</td>
      <td class="tdbg"><input name="t2" type="text" class="input" value="<%=Sdcms_Create_GoogleMap(0)%>" />��<span>Ϊ0ʱ��ʾȫ��</span></td>
    </tr>
	<tr>
      <td height="25" align="center" class="tdbg">GoogleƵ�ʣ�</td>
      <td class="tdbg"><select name="t3"><option value="always" <%=IIF(Sdcms_Create_GoogleMap(1)="always","selected","")%>>Always</option><option value="hourly" <%=IIF(Sdcms_Create_GoogleMap(1)="hourly","selected","")%>>Hourly</option><option value="daily" <%=IIF(Sdcms_Create_GoogleMap(1)="daily","selected","")%>>Daily</option><option value="weekly" <%=IIF(Sdcms_Create_GoogleMap(1)="weekly","selected","")%>>Weekly</option><option value="monthly" <%=IIF(Sdcms_Create_GoogleMap(1)="monthly","selected","")%>>Monthly</option><option value="yearly" <%=IIF(Sdcms_Create_GoogleMap(1)="yearly","selected","")%>>Yearly</option></select></td>
    </tr>
	<tr>
      <td height="25" align="center" class="tdbg">Google����Ȩ��</td>
      <td class="tdbg"><input name="t4" type="text" class="input" value="<%=Sdcms_Create_GoogleMap(2)%>" />��<span>0-1֮�������</span></td>
    </tr>
    <tr>
      <td height="25" colspan="2" class="tdbg"><input type="hidden" name="act" value="5" /><input name="Submit" type="submit" class="bnt" value="��ʼ����" /></td>
    </tr>
	</form>
  </table>
<%End Sub:Sub Create_Msg%>
 <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <tr>
      <td class="tdbg01">���ɽ���</td>
    </tr>
    <tr>
      <td height="25" class="tdbg" style="padding:10px;"><iframe frameborder="0" name="Loadinfo" id="Loadinfo" width="100%" height="150"></iframe></td>
    </tr>
  </table>
<%Db_Run:End Sub:Sub NoHtml%>
 <table border="0" align="center" cellpadding="3" cellspacing="1" class="table_b mag_t">
    <tr>
      <td class="tdbg01">������ʾ</td>
    </tr>
    <tr>
      <td height="25" class="tdbg" style="padding:10px;">��ģʽ������Ҫ���ɣ�</td>
    </tr>
  </table>
<%End Sub%>