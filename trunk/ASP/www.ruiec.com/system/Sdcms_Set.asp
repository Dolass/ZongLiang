<!--#include file="sdcms_check.asp"-->
<!--#include file="../Inc/AspJpeg.asp"-->
<%
Dim Sdcms,title,Sd_Table,Action,ID
ID=IsNum(Trim(Request.QueryString("ID")),0)
Action=Lcase(Trim(Request.QueryString("Action")))
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Sdcms.Check_lever 1
Set sdcms=Nothing
title="ϵͳ����"
Sd_Table="sd_const"
Sdcms_Head
%>
<div class="sdcms_notice"><span>���������</span><a href="?">ϵͳ����</a><!--������<a href="?">��Ա����</a>--></div>
<br>
<ul id="sdcms_sub_title">
	<li class="sub"><a href="javascript:void(0)" onClick="selectTag('tagContent0',this)"><%=title%></a></li>
	<li class="unsub"><a href="javascript:void(0)" onClick="selectTag('tagContent1',this)">��������</a></li>
	<li class="unsub"><a href="javascript:void(0)" onClick="selectTag('tagContent2',this)">��������</a></li>
	<li class="unsub"><a href="javascript:void(0)" onClick="selectTag('tagContent3',this)">�ϴ�����</a></li>
</ul>
<div id="sdcms_right_b">
<%
Select Case Action
	Case "save":save
	Case Else:main
End Select
Db_Run
Closedb
Sub Main
	Dim Rs,i,all_set,FontArr
	Set Rs=Conn.Execute("Select id,webname,weburl,webkey,webdec From "&Sd_Table&" Where Id=1")
	DbQuery=DbQuery+1
	IF Rs.eof then
		Echo "����Ƿ��ύ����":Exit Sub
	End IF
	Echo Check_Add
%><form name="add" method="post" action="?action=save" onSubmit="return checkadd()">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" id="tagContent0">
    <tr>
      <td width="150" align="center" class="tdbg">��վ���ƣ�</td>
      <td class="tdbg"><input name="t0" type="text" class="input" value="<%=rs(1)%>" size="30"></td>
    </tr>
    <tr class="tdbg">
      <td align="center">��վ������</td>
      <td><input name="t1" type="text" class="input" value="<%=rs(2)%>"  size="30">��<span>��ʽΪ��http://www.sdcms.cn������Ҫ�ӡ�/��</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">����ģʽ��</td>
      <td><input name="t23" type="radio" value="0" <%=IIF(Sdcms_Mode=0,"Checked","")%> id="t23_0" onclick="$('#t3')[0].value='.asp';$('#modetip')[0].style.display='none'"><label for="t23_0">��̬</label> <input name="t23" type="radio"  value="1" <%=IIF(Sdcms_Mode=1,"Checked","")%> id="t23_1"  onclick="$('#t3')[0].value='.html';$('#modetip')[0].style.display='inline'" <%if len(request.ServerVariables("HTTP_X_REwrite_url"))=0 then%>disabled="disabled"<%end if%>><label for="t23_1">α��̬</label> <input name="t23" type="radio"  value="2" <%=IIF(Sdcms_Mode=2,"Checked","")%> id="t23_2" onclick="$('#t3')[0].value='.html';$('#modetip')[0].style.display='none'"><label for="t23_2">��̬</label>��<span id="modetip" class="<%=IIF(Sdcms_Mode<>1,"dis","")%>">α��̬ģʽ��Ҫ�ռ�֧��Rewrite���</span></td>
    </tr>
    <tr class="tdbg">
      <td align="center">����ģʽ��</td>
      <td><input name="t26" type="radio" value="true" <%=IIF(blogmode,"Checked","")%> id="t26_0"  ><label for="t26_0">����</label> <input name="t26" type="radio"  value="false" <%=IIF(not(blogmode),"Checked","")%> id="t26_1"><label for="t26_1">�ر�</label></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ļ���׺��</td>
      <td><select name="t3" id="t3">
	    <option selected>��ѡ���ļ���׺</option>
		<option value=".asp" <%=IIF(Lcase(Sdcms_filetxt)=".asp","Selected","")%>>.Asp</option>
        <option value=".html" <%=IIF(Lcase(Sdcms_filetxt)=".html","Selected","")%>>.Html</option>
      </select>��<span>��������˺�׺��Ҫ����ȫվ�Ż���Ч</span>      </td>
    </tr>
    <tr class="tdbg">
      <td align="center">��ȡ���ȣ�</td>
      <td><input name="t22" type="text" class="input" value="<%=Sdcms_length%>" size="30">��<span>��Ϣ�����Զ���ȡ���ȣ����飺200-500</span></td>
	</tr>
	<tr class="tdbg">
      <td align="center">ʱ��ѡ��</td>
      <td><select name="t25" id="t25" style="width:400px;">
	    <option selected>��ѡ�����������ʱ��</option>
		<option value="-20"<%=IIF(Sdcms_TimeZone=-20," selected","")%>>(<%=Dateadd("h",-20,Now())%>) ���������ˡ��ϼ���</option>
		<option value="-19"<%=IIF(Sdcms_TimeZone=-19," selected","")%>>(<%=Dateadd("h",-19,Now())%>) ��;������Ħ��Ⱥ��</option>
		<option value="-18"<%=IIF(Sdcms_TimeZone=-18," selected","")%>>(<%=Dateadd("h",-18,Now())%>) ������</option>
		<option value="-17"<%=IIF(Sdcms_TimeZone=-17," selected","")%>>(<%=Dateadd("h",-17,Now())%>) ����˹��</option>
		<option value="-16"<%=IIF(Sdcms_TimeZone=-16," selected","")%>>(<%=Dateadd("h",-16,Now())%>) ̫ƽ��ʱ�䣨�����ͼ��ô󣩣��ٻ���</option>
		<option value="-15"<%=IIF(Sdcms_TimeZone=-15," selected","")%>>(<%=Dateadd("h",-15,Now())%>) ɽ��ʱ��(����)������ɣ��</option>
		<option value="-14"<%=IIF(Sdcms_TimeZone=-14," selected","")%>>(<%=Dateadd("h",-14,Now())%>) �в�ʱ�䣨�����ͼ��ô󣩣��ع����Ӷ��ͣ���˹������ʡ��ī����ǡ�����������</option>
		<option value="-13"<%=IIF(Sdcms_TimeZone=-13," selected","")%>>(<%=Dateadd("h",-13,Now())%>) ����ʱ�䣨�����ͼ��ô󣩡��������������</option>
		<option value="-12"<%=IIF(Sdcms_TimeZone=-12," selected","")%>>(<%=Dateadd("h",-12,Now())%>) ������ʱ�䣨���ô�ί������������˹</option>
		<option value="-11.5"<%=IIF(Sdcms_TimeZone=-11.5," selected","")%>>(<%=Dateadd("h",-11.5,Now())%>) �µ�(���ô󶫰�) Ŧ����</option>
		<option value="-11"<%=IIF(Sdcms_TimeZone=-11," selected","")%>>(<%=Dateadd("h",-11,Now())%>) �������� �������� ��³ŵ˹����˹�����γ�</option>
		<option value="-10"<%=IIF(Sdcms_TimeZone=-10," selected","")%>>(<%=Dateadd("h",-10,Now())%>) �������в�</option>
		<option value="-9"<%=IIF(Sdcms_TimeZone=-9," selected","")%>>(<%=Dateadd("h",-9,Now())%>) ���ٶ�Ⱥ������ý�Ⱥ��</option>
		<option value="-8"<%=IIF(Sdcms_TimeZone=-8," selected","")%>>(<%=Dateadd("h",-8,Now())%>) �������α�׼ʱ�� ���׶أ������֣�����������˹��������������������ά�ǣ�Ӣ������</option>
		<option value="-7"<%=IIF(Sdcms_TimeZone=-7," selected","")%>>(<%=Dateadd("h",-7,Now())%>) ��������ʿ������ά�ǣ�˹�工�ˣ�˹�������ǣ��ݿˣ����������������������ѣ�����٣��������ǣ����������޵���</option>
		<option value="-6"<%=IIF(Sdcms_TimeZone=-6," selected","")%>>(<%=Dateadd("h",-6,Now())%>) ������˹�أ������ף��Ϸǣ����������ǣ��������ն���������ӣ����֣���˹�ˣ���ɫ��</option>
		<option value="-5"<%=IIF(Sdcms_TimeZone=-5," selected","")%>>(<%=Dateadd("h",-5,Now())%>) ɳ�ڵذ�����������˹ ���ŵã�Ī˹�ƣ�ʥ�˵ñ��������Ӹ��գ����ޱ�</option>
		<option value="-5.5"<%=IIF(Sdcms_TimeZone=-5.5," selected","")%>>(<%=Dateadd("h",-5.5,Now())%>) ����</option>
		<option value="-4"<%=IIF(Sdcms_TimeZone=-4," selected","")%>>(<%=Dateadd("h",-4,Now())%>) �������ȣ���˹���أ��Ϳ⡢�ڱ���˹���������(��������)��Ī˹����������˹(�����ǹ���)��������</option>
		<option value="-4.5"<%=IIF(Sdcms_TimeZone=-4.5," selected","")%>>(<%=Dateadd("h",-4.5,Now())%>) ������</option>
		<option value="-3"<%=IIF(Sdcms_TimeZone=-3," selected","")%>>(<%=Dateadd("h",-3,Now())%>) ���� : Ҷ�����ձ��������桢��ʲ�ɡ���˹����͵¡������桢�������ֱ�</option>
		<option value="-3.5"<%=IIF(Sdcms_TimeZone=-3.5," selected","")%>>(<%=Dateadd("h",-3.5,Now())%>) ӡ�� : ���򣬼Ӷ����������˹���µ���</option>
		<option value="-2"<%=IIF(Sdcms_TimeZone=-2," selected","")%>>(<%=Dateadd("h",-2,Now())%>) ���� : ����ľͼ�������¡������ᡢ�￨</option>
		<option value="-1"<%=IIF(Sdcms_TimeZone=-1," selected","")%>>(<%=Dateadd("h",-1,Now())%>) ���ȣ����ڣ��żӴ�</option>
		<option value="0"<%=IIF(Sdcms_TimeZone=0," selected","")%>>(<%=Dateadd("h",0,Now())%>) �й� : ���������죬���ݣ��Ϻ�����ۣ�̨�����¼���</option>
		<option value="1"<%=IIF(Sdcms_TimeZone=1," selected","")%>>(<%=Dateadd("h",1,Now())%>) ƽ�������ǣ����������棬���ϣ��ſ�Ŀ�</option>
		<option value="1.5"<%=IIF(Sdcms_TimeZone=1.5," selected","")%>>(<%=Dateadd("h",1.5,Now())%>) �����в�</option>
		<option value="2"<%=IIF(Sdcms_TimeZone=2," selected","")%>>(<%=Dateadd("h",2,Now())%>) ��̫ƽ�� : ϯ���ᡢ��˹÷���ǡ��ص���Ī���ȱȸۣ������أ���������Ϥ��</option>
		<option value="3"<%=IIF(Sdcms_TimeZone=3," selected","")%>>(<%=Dateadd("h",3,Now())%>) ̫ƽ���в� : ��ӵ���������Ⱥ�����¿��������</option>
		<option value="4"<%=IIF(Sdcms_TimeZone=4," selected","")%>>(<%=Dateadd("h",4,Now())%>) Ŧ���� : ����١�쳼ã�����Ӱ뵺�����ܶ�Ⱥ��</option>
      </select>��<span>�������������</span>      </td>
    </tr>
	<tr class="tdbg">
      <td align="center">�Ƿ񻺴棺</td>
      <td><input name="t6" type="radio" value="True" onclick="$('#cookies')[0].style.display='block';$('#cookies_date')[0].style.display='block';" <%=IIF(Sdcms_Cache=True,"Checked","")%> id="t6_0"><label for="t6_0">����</label> <input name="t6" type="radio"  value="False" onclick="$('#cookies')[0].style.display='none';$('#cookies_date')[0].style.display='none';" <%=IIF(Sdcms_Cache=False,"Checked","")%> id="t6_1"><label for="t6_1">�ر�</label></td>
    </tr>
	<tr class="tdbg<%IF Not Sdcms_Cache Then%> dis<%End IF%>" id="cookies">
      <td align="center">����ǰ׺��</td>
      <td><input name="t12" type="text" class="input" value="<%=Sdcms_cookies%>"  size="30">��<span>����ж��վ�㣬�����ò�ͬ��ֵ</span></td>
    </tr>
	<tr class="tdbg<%IF Not Sdcms_Cache Then%> dis<%End IF%>" id="cookies_date">
      <td align="center">����ʱ�䣺</td>
      <td><input name="t24" type="text" class="input" value="<%=Sdcms_CacheDate%>" size="30">��<span>��λ����</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">������־��</td>
      <td><input name="t18" type="radio" value="True" <%=IIF(Sdcms_AdminLog=true,"Checked","")%> id="t18_0"><label for="t18_0">����</label> <input name="t18" type="radio"  value="False" <%=IIF(Sdcms_AdminLog=false,"Checked","")%> id="t18_1"><label for="t18_1">�ر�</label> (<span>�ر���ֻ��¼��½��־</span>)</td>
    </tr>
	<tr class="tdbg">
      <td align="center" width="14%">���ۿ��أ�</td>
      <td><input name="t7" type="radio" value="0" <%=IIF(Sdcms_comment_pass=0,"Checked","")%> id="t7_0"><label for="t7_0">�ر�</label> <input name="t7" type="radio"  value="1" <%=IIF(Sdcms_comment_pass=1,"Checked","")%> id="t7_1"><label for="t7_1">����</label></td>
    </tr>
	<tr class="tdbg">
      <td align="center">������ˣ�</td>
      <td><input name="t8" type="radio" value="0" <%=IIF(Sdcms_comment_ispass=0,"Checked","")%> id="t8_0"><label for="t8_0">�ر�</label> <input name="t8" type="radio"  value="1" <%=IIF(Sdcms_comment_ispass=1,"Checked","")%> id="t8_1"><label for="t8_1">����</label></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�� �� �֣�</td>
      <td><textarea name="t4"  rows="3"  class="inputs" id="t4"><%=Content_Encode(Rs(3))%></textarea></td>
    </tr>
	<tr class="tdbg">
      <td align="center">��վ������</td>
      <td><textarea name="t5"  rows="3"  class="inputs" id="t5"><%=Content_Encode(Rs(4))%></textarea></td>
	</tr>
	</table>

	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1"  class="dis" id="tagContent1">
	<tr class="tdbg">
      <td width="120" align="center">HTML��ǩ���ˣ�</td>
      <td><textarea name="t9"  rows="6" class="inputs"><%=Content_Encode(Sdcms_badhtml)%></textarea><span>������á�|���ָ�</span></td>
	</tr>
	<tr class="tdbg">
      <td align="center">HTML�¼����ˣ�</td>
      <td><textarea name="t10" rows="6" class="inputs"><%=Content_Encode(Sdcms_badEvent)%></textarea><span>������á�|���ָ�</span></td>
	</tr>
	<tr class="tdbg">
      <td align="center">�໰���ˣ�</td>
      <td><textarea name="t11" rows="6" class="inputs"><%=Content_Encode(Sdcms_badtext)%></textarea><span>������á�|���ָ�</span></td>
	</tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="dis" id="tagContent2">
	<tr class="tdbg">
      <td colspan="2">�� <b>����Ϊƫ�����ã�����Ϊȫ�ֲ�����</b></td>
    </tr>
	<tr class="tdbg">
      <td width="120" align="center">����Ŀ¼��</td>
      <td><input name="t17" type="text" class="input" value="<%=Sdcms_htmdir%>"  size="30">��<span>��Ŀ¼Ϊ��"��Ҳ����ָ��ΪĳһĿ¼����ʽΪ��"sdcms/"</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ļ����ƣ�</td>
      <td><select name="t16"><option value="{ID}" <%=IIF(Sdcms_filename="{ID}","selected","")%>>�Զ����</option><option value="{YYMMDDID}" <%=IIF(Sdcms_filename="{YYMMDDID}","selected","")%>>��+��+��+���</option><option value="{PinYin}" <%=IIF(Sdcms_filename="{PinYin}","selected","")%>>����ƴ��</option></select></td>
    </tr>
	<tr class="tdbg">
      <td align="center">��Ϣ����</td>
      <td><%all_set="������ҳ|���ɷ���|������Ϣ":all_set=Split(all_set,"|"):For I=0 To Ubound(all_set)%><input name="t15" type="checkbox" value="<%=I%>" <%=IIF(Instr(", "&Sdcms_Create_Set&", ",", "&I&", ")>0,"Checked","")%> id="<%=I%>"><label for="<%=I%>"><%=all_set(i)%></label>
	  <%next%></td>
    </tr>
	<tr class="tdbg">
      <td colspan="2"><b>Google��ͼѡ�</b></td>
      </tr>
	  <tr>
      <td align="center" class="tdbg">����������</td>
      <td class="tdbg"><input name="t21" type="text" class="input" value="<%=Sdcms_Create_GoogleMap(0)%>" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" />��<span>Ϊ0ʱ��ʾȫ��</span></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">Ƶ�����ʣ�</td>
      <td class="tdbg"><select name="t21"><option value="always" <%=IIF(Sdcms_Create_GoogleMap(1)="always","Selected","")%>>Always</option><option value="hourly" <%=IIF(Sdcms_Create_GoogleMap(1)="hourly","Selected","")%>>Hourly</option><option value="daily" <%=IIF(Sdcms_Create_GoogleMap(1)="daily","Selected","")%>>Daily</option><option value="weekly"  <%=IIF(Sdcms_Create_GoogleMap(1)="weekly","Selected","")%>>Weekly</option><option value="monthly"  <%=IIF(Sdcms_Create_GoogleMap(1)="monthly","Selected","")%>>Monthly</option><option value="yearly"  <%=IIF(Sdcms_Create_GoogleMap(1)="yearly","Selected","")%>>Yearly</option></select></td>
    </tr>
	<tr>
      <td align="center" class="tdbg">�� �� Ȩ��</td>
      <td class="tdbg"><input name="t21" type="text" class="input" value="<%=Sdcms_Create_GoogleMap(2)%>" />��<span>0-1֮�������</span></td>
    </tr> 
	</table>
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="dis" id="tagContent3">
	<tr class="tdbg">
      <td align="center" width="120">�ϴ�Ŀ¼��</td>
      <td><input name="t14" type="text" class="input" value="<%=Sdcms_upfiledir%>"  size="40">��<span>����Ĭ�����ú���ܻ�Ӱ����ǰ����Ϣ</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ļ����ͣ�</td>
      <td><input name="t19" type="text" class="input" value="<%=Sdcms_upfiletype%>"  size="40">��<span>�����ϴ����ļ����ͣ�������á�|����</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�ļ����ƣ�</td>
      <td><input name="t20" type="text" class="input" value="<%=Sdcms_upfileMaxSize%>" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"  size="40">��<span>�����ϴ����ļ����ֵ����λ��KB</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">ͼ�������</td>
      <td><select name="f0" size="1" id="f0"><option value="false" <%=IIF(Sdcms_Jpeg_t0=False,"Selected","")%>>�ر�</option><option value="True" <%=IIF(Sdcms_Jpeg_t0=True,"Selected","")%>>AspJpeg<%=IIF(IsObjInstalled("Persits.Jpeg"),"��","��")%></option></select></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�� �� ͼ��</td>
      <td>��ȣ�<input name="f13" type="text" value="<%=Sdcms_Jpeg_t13%>"  onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" class="input" size="5" /> �߶ȣ�<input name="f14" type="text" value="<%=Sdcms_Jpeg_t14%>" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" class="input" size="5" />��<span>Ϊ0ʱ����������С</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">�㡡������</td>
      <td><select name="f15" size="1" onchange="<%For i=0 To 2%>$('#sf<%=i%>')[0].style.display='none';<%Next%>$('#sf'+this.value)[0].style.display='inline';"><option value="0" <%=IIF(Sdcms_Jpeg_t15=0,"Selected","")%>>���淨</option><option value="1" <%=IIF(Sdcms_Jpeg_t15=1,"Selected","")%>>�ü���</option><option value="2" <%=IIF(Sdcms_Jpeg_t15=2,"Selected","")%>>���䷨</option></select>��<span id="sf0" <%=IIF(Sdcms_Jpeg_t15<>0,"class=""dis""","")%>>���淨����Ⱥ͸߶ȶ�����0ʱ��ֱ����С��ָ����С������һ��Ϊ0ʱ����������С</span><span id="sf1" <%=IIF(Sdcms_Jpeg_t15<>1,"class=""dis""","")%>>�ü�������Ⱥ͸߶ȶ�����0ʱ���Ȱ���ѱ�����С�ٲü���ָ����С������һ��Ϊ0ʱ����������С</span><span id="sf2" <%=IIF(Sdcms_Jpeg_t15<>2,"class=""dis""","")%>>���䷨����ָ����С�ı���ͼ�ϸ����ϰ���ѱ�����С��ͼƬ</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">ˮӡ���ͣ�</td>
      <td><input name="f1" type="radio" value="True" <%=IIF(Sdcms_Jpeg_t1=true,"checked","")%> onclick="<%for i=1 to 5%>$('#jpeg0<%=i%>')[0].style.display='block';<%next%><%For i=6 to 8%>$('#jpeg0<%=i%>')[0].style.display='none';<%next%>" id="f1_0"><label for="f1_0">����ˮӡ</label> <input name="f1" type="radio" value="False"  <%=IIF(Sdcms_Jpeg_t1=false,"checked","")%>  onclick="<%for i=1 to 5%>$('#jpeg0<%=i%>')[0].style.display='none';<%next%><%for i=6 to 8%>$('#jpeg0<%=i%>')[0].style.display='block';<%next%>" id="f1_1"><label for="f1_1">ͼƬˮӡ</label></td>
    </tr>
	<tr class="tdbg">
      <td align="center">ˮӡ������</td>
      <td><input name="f2" type="text" value="<%=Sdcms_Jpeg_t2%>" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" class="input" />��<span>0-100 Ϊ100ʱ�������</span></td>
    </tr>
	<tr class="tdbg <%IF Not Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg01">
      <td align="center">ˮӡ���֣�</td>
      <td><input name="f3" type="text" value="<%=Sdcms_Jpeg_t3%>" class="input" /></td>
    </tr>
	<tr class="tdbg <%IF Not Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg02">
      <td align="center">�������壺</td>
      <td><select name="f4" size="1">
	  <%FontArr=Split("����|����_GB2312|����|����|��Բ|Arial|Verdana","|"):For I=0 To Ubound(FontArr)%>
	  <option value="<%=FontArr(I)%>" <%=IIF(Sdcms_Jpeg_t4=FontArr(I),"selected","")%>><%=FontArr(I)%></option>
	  <%Next%>
	  </select></td>
    </tr>
	<tr class="tdbg <%IF Not Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg03">
      <td align="center">�����С��</td>
      <td><input name="f5" type="text" value="<%=Sdcms_Jpeg_t5%>" onKeypress="if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;" class="input" />��<span>��λ��Px</span></td>
    </tr>
	<tr class="tdbg <%IF Not Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg04">
      <td align="center">������ɫ��</td>
      <td><input name="f6" type="text" value="<%=Replace(Sdcms_Jpeg_t6,"&H","#")%>" maxlength="7" class="input" />��<span>�磺#000000</span></td>
    </tr>
	<tr class="tdbg <%IF Not Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg05">
      <td align="center">�Ƿ�Ӵ֣�</td>
      <td><input name="f7" type="radio" value="1" <%=IIF(Sdcms_Jpeg_t7=1,"checked","")%> id="f7_0"><label for="f7_0">��</label> <input name="f7" type="radio" value="0" <%=IIF(Sdcms_Jpeg_t7=0,"checked","")%> id="f7_1"><label for="f7_1">��</label></td>
    </tr>
	<tr class="tdbg <%IF Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg06">
      <td align="center">ˮӡͼƬ��</td>
      <td><input name="f8" id="f8" type="text" value="<%=Sdcms_Jpeg_t8%>" class="input" /><%admin_upfile 1,"100%","20","f8","UpLoadIframe",0,0%></td>
    </tr>
	<tr class="tdbg <%IF Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg07">
      <td align="center" >͸ �� �ȣ�</td>
      <td><input name="f9" type="text" value="<%=Sdcms_Jpeg_t9%>" class="input" />��<span>0-1֮�������</span></td>
    </tr>
	<tr class="tdbg <%IF Sdcms_Jpeg_t1 Then%>dis<%End IF%>" id="jpeg08">
      <td align="center">ͼƬ��ɫ��</td>
      <td><input name="f10" type="text" value="<%=Replace(Sdcms_Jpeg_t10,"&H","#")%>" class="input" />��<span>�磺#000000��Ϊ��ʱ��ȥ��ˮӡͼ��ɫ</span></td>
    </tr>
	<tr class="tdbg">
      <td align="center">������㣺</td>
      <td><select name="f11" size="1">
	  <option value="0" <%=IIF(Sdcms_Jpeg_t11=0,"selected","")%>>����</option>
	  <option value="1" <%=IIF(Sdcms_Jpeg_t11=1,"selected","")%>>����</option>
	  <option value="2" <%=IIF(Sdcms_Jpeg_t11=2,"selected","")%>>����</option>
	  <option value="3" <%=IIF(Sdcms_Jpeg_t11=3,"selected","")%>>����</option>
	  <option value="4" <%=IIF(Sdcms_Jpeg_t11=4,"selected","")%>>����</option>
	  </select></td>
    </tr>
	<tr class="tdbg">
      <td align="center">����λ�ã�</td>
      <td>X�᣺<input name="f12_0" type="text" class="input" onKeypress="if(event.keyCode<45||event.keyCode>57)event.returnValue=false;" value="<%=Sdcms_Jpeg_t12(0)%>" maxlength="5" />��Y�᣺<input name="f12_1" type="text" value="<%=Sdcms_Jpeg_t12(1)%>" class="input" onKeypress="if(event.keyCode<45||event.keyCode>57)event.returnValue=false;" maxlength="5" />��<span>ֻ��������</span></td>
    </tr>
	</table>
	<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1">
    <tr class="tdbg">
      <td>���������������������� <input name="Submit" type="submit" class="bnt" value="��������"></td>
    </tr>
	</table>
	</form>
<%
	Rs.Close
	Set Rs=Nothing
End Sub

Sub Save
	Dim t0,t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15,t16,t17,t18,t19,t20,t21,t22,t23,t24,t25,t26
	Dim f0,f1,f2,f3,f4,f5,f6,f7,f8,f9,f10,f11,f12_0,f12_1,f13,f14,f15
	Dim c0,c1,c2,c3,c4,c5,c6,c7
	Dim Old_DataFile,Old_DataName
	t0=FilterText(Trim(Request.Form("t0")),1)
	t1=FilterText(Trim(Request.Form("t1")),0)
	t3=Trim(Request.Form("t3"))'�ļ���׺
	t4=FilterHtml(Trim(Request.Form("t4")))'�ؼ���
	t5=FilterHtml(Trim(Request.Form("t5")))'����
	t6=FilterText(Trim(Request.Form("t6")),1)
	t7=FilterText(Trim(Request.Form("t7")),1)
	t8=FilterText(Trim(Request.Form("t8")),1)
	t9=FilterText(Trim(Request.Form("t9")),0)'���˿�ʼ
	t10=FilterText(Trim(Request.Form("t10")),0)
	t11=FilterText(Trim(Request.Form("t11")),0)'���˽���
	t12=FilterText(Trim(Request.Form("t12")),1)
	t14=FilterText(Trim(Request.Form("t14")),0)'�ϴ�Ŀ¼
	t15=FilterText(Trim(Request.Form("t15")),1)
	t16=FilterText(Trim(Request.Form("t16")),0)
	t17=FilterText(Trim(Request.Form("t17")),0)'����Ŀ¼
	t18=FilterText(Trim(Request.Form("t18")),0)
	t19=FilterText(Trim(Request.Form("t19")),0)'�ļ�����
	t20=IsNum(Trim(Request.Form("t20")),200)
	t21=Replace(Trim(Request.Form("t21"))," ","")
	t22=IsNum(Trim(Request.Form("t22")),200)
	t23=FilterText(Trim(Request.Form("t23")),1)
	t24=IsNum(Trim(Request.Form("t24")),60)
	t25=IsNum(Trim(Request.Form("t25")),0)
	t26=FilterText(Trim(Request.Form("t26")),0)'�ļ�����
	
	f0=FilterText(Trim(Request.Form("f0")),1)
	f1=FilterText(Trim(Request.Form("f1")),1)
	f2=IsNum(Trim(Request.Form("f2")),100)
	f3=FilterText(Trim(Request.Form("f3")),0)'ˮӡ����
	f4=FilterText(Trim(Request.Form("f4")),1)
	f5=IsNum(Trim(Request.Form("f5")),12)
	f6=FilterText(Trim(Request.Form("f6")),1)
	f7=FilterText(Trim(Request.Form("f7")),1)
	f8=FilterText(Trim(Request.Form("f8")),0)
	f9=FilterText(Trim(Request.Form("f9")),1)
	f10=FilterText(Trim(Request.Form("f10")),1)
	f11=FilterText(Trim(Request.Form("f11")),1)
	f12_0=IsNum(Trim(Request.Form("f12_0")),0)
	f12_1=IsNum(Trim(Request.Form("f12_1")),0)
	f13=IsNum(Trim(Request.Form("f13")),150)
	f14=IsNum(Trim(Request.Form("f14")),113)
	f15=FilterText(Trim(Request.Form("f15")),1)
	f6=Replace(f6,"#","&H")
	IF Len(f10)>1 Then f10=Replace(f10,"#","&H") Else f10=""
	
	If t17<>"" And Right(t17,"1")<>"/" Then t17=t17&"/"
	t17=re(t17,".","")
	Select case Lcase(t3)
		Case ".asp",".htm",".html",".shtml"
		Case Else:t3=".html"
	End Select
	
	Select Case t13
		Case "0","1"
		Case Else:t13=0
	End Select
	
	IF t23=0 Then t3=".asp" Else IF t23=2 And t3=".asp" Then t3=".html"
	if t23=1 then t3=".html"

	t9=check_event(t9,"|",""):t10=check_event(t10,"|",""):t11=check_event(t11,"|",""):t19=check_event(t19,"|",3)
	Dim badfilename,I
	badfilename=Split("asp|aspx|jsp|asa|html|htm|js|vbs|exe|cer|cdx|htw|ida|idq|shtm|shtml|stm|printer|cgi|php|php4|cfm|ashx","|")
	For I=0 To Ubound(badfilename)
		t19=Replace(t19,badfilename(I),"")
	Next
	IF Right(t19,1)="|" Then t19=Left(t19,Len(t19)-1)
	t19=Replace(t19,".","")
	Dim Rs,Sql
	Set Rs=server.CreateObject("Adodb.RecordSet")
	Sql="Select WebName,WebUrl,WebKey,WebDec From "&Sd_Table&" Where Id=1"
	Rs.Open Sql,Conn,1,3
	Rs.Update
		Rs(0)=left(t0,255)
		Rs(1)=left(t1,255)
		Rs(2)=left(t4,255)
		Rs(3)=left(t5,255)
	Rs.Update
	Rs.Close
	Set Rs=Nothing
	'������վ����
	
	ReName_Folder "../"&Sdcms_Upfiledir,"../"&t14'����������Ŀ¼
	Dim Change_Mode
	Change_Mode=False
	IF Sdcms_Mode<2 Then
		IF t23=2 Then
			Change_Mode=True
		End IF
	End IF
	IF Sdcms_Mode=2 Then
		IF t23<2 Then
			Change_Mode=True
		End IF
	End IF
	
	if t23=1 then
		'����httpd.ini
		dim httpd
		httpd=""
		httpd=httpd&"[ISAPI_Rewrite]"&vbcrlf&vbcrlf

		httpd=httpd&"# 3600 = 1 hour"&vbcrlf
		httpd=httpd&"#CacheClockRate 3600"&vbcrlf&vbcrlf
		
		httpd=httpd&"RepeatLimit 32"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д��ҳ"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Page/(.*)_(\d*)\.html "&Sdcms_Root&"page/\?id=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Page/(.*)\.html "&Sdcms_Root&"page/\?id=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дר��"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Topic/List_(.*)_(\d*)\.html "&Sdcms_Root&"Topic/\List.asp\?ID=$1&Page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Topic/List_(.*)\.html "&Sdcms_Root&"Topic/\List.asp\?ID=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Topic/Index_(.*)\.html "&Sdcms_Root&"Topic/\?page=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"Topic/Index\.html "&Sdcms_Root&"Topic/\index.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дTags"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"tags/List_(.*)\.html "&Sdcms_Root&"tags/\List.asp\?page=$1 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"tags/List\.html(.*) "&Sdcms_Root&"tags/\list.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"RewriteRule "&Sdcms_Root&"tags/(.*)_(\d*)\.html "&Sdcms_Root&"tags/\?/$1/$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"tags/(.*)\.html "&Sdcms_Root&"tags/\?/$1/ [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дindex.asp"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"index\.html(.*) "&Sdcms_Root&"index.asp [N,I]"&vbcrlf&vbcrlf
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"index_(\d*)\.html "&Sdcms_Root&"index.asp\?page=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��дsitemap.asp"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"sitemap\.html(.*) "&Sdcms_Root&"sitemap.asp [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"html/(.*)/(\d*)_(\d*)\.html "&Sdcms_Root&"Info/\View.asp\?ID=$2&page=$3 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"html/(.*)/(\d*)\.html "&Sdcms_Root&"Info/\View.asp\?ID=$2 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д�б�"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"html/(.*)/(\d*) "&Sdcms_Root&"Info/\?cname=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"html/(.*)/ "&Sdcms_Root&"Info/\?cname=$1 [N,I]"&vbcrlf&vbcrlf
		
		httpd=httpd&"#��д����"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"date-(.*)_(\d*) "&Sdcms_Root&"date.asp\?c=$1&page=$2 [N,I]"&vbcrlf
		httpd=httpd&"RewriteRule "&Sdcms_Root&"date-(.*) "&Sdcms_Root&"date.asp\?c=$1 [N,I]"
		Savefile "/","httpd.ini",httpd
	else
		'ɾ��httpd.ini��httpd.parse.errors
		Del_File "/httpd.ini"
		Del_File "/httpd.parse.errors"
	end if
	
	Dim Config
	Config=""
	Config=Config&"<"
	Config=Config&"%"& vbcrlf
	Config=Config&"'��վ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_WebName:Sdcms_WebName="""&t0&""""&vbcrlf
	Config=Config&"'��վ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_WebUrl:Sdcms_WebUrl="""&t1&""""&vbcrlf
	Config=Config&"'ϵͳĿ¼����Ŀ¼Ϊ��""/"",����Ŀ¼��ʽΪ��""/sdcms/"""& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Root:Sdcms_Root="""&Sdcms_Root&""""&vbcrlf
	Config=Config&"'����ģʽ,0Ϊ��̬Ĭ�ϣ�1Ϊα��̬��2Ϊ��̬"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Mode:Sdcms_Mode="&t23&""&vbcrlf
	Config=Config&"'����ģʽ"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim blogmode:blogmode="&t26&""&vbcrlf
	Config=Config&"'����Ŀ¼����Ŀ¼Ϊ��,Ҳ����ָ��ΪĳһĿ¼����ʽΪ��""sdcms/"",�����ԣ�""/""����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_HtmDir:Sdcms_HtmDir="""&t17&""""&vbcrlf
	Config=Config&"'ʱ������,�벻Ҫ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_TimeZone:Sdcms_TimeZone="&t25&vbcrlf
	Config=Config&"'�Ƿ���������־"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_AdminLog:Sdcms_AdminLog="&t18&vbcrlf
	Config=Config&"'�Ƿ�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cache:Sdcms_Cache="&t6&vbcrlf
	Config=Config&"'����ǰ׺,����ж��վ�㣬�����ò�ͬ��ֵ"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Cookies:Sdcms_Cookies="""&t12&""""&vbcrlf
	Config=Config&"'����ʱ�䣬��λΪ��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_CacheDate:Sdcms_CacheDate="&t24&""&vbcrlf
	Config=Config&"'ϵͳ�ļ��ĺ�׺�������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileTxt:Sdcms_FileTxt="""&t3&""""&vbcrlf
	Config=Config&"'ϵͳģ��Ŀ¼,һ�㲻�����ֶ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Skins_Root:Sdcms_Skins_Root="""&Sdcms_Skins_Root&""""&vbcrlf
	Config=Config&"'ϵͳ�����ļ���Ĭ���ļ��������鲻Ҫ�Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_FileName:Sdcms_FileName="""&t16&""""&vbcrlf
	Config=Config&"'����Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileDir:Sdcms_UpfileDir="""&t14&""""&vbcrlf
	Config=Config&"'�����ϴ��ļ�����,�������|��"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_UpfileType:Sdcms_UpfileType="""&t19&""""&vbcrlf
	Config=Config&"'�����ϴ����ļ����ֵ����λ��KB"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_upfileMaxSize:Sdcms_upfileMaxSize="&t20&vbcrlf
	Config=Config&"'����ƫ������,ϵͳ�Զ����ɣ���������Ķ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Create_Set:Sdcms_Create_Set="""&Re(t15,"&nbsp;"," ")&""""&vbcrlf
	Config=Config&"'����GOOGLE��ͼ�Ĳ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Create_GoogleMap:Sdcms_Create_GoogleMap=Split("""&t21&""","","")"&vbcrlf
	Config=Config&"'���ۿ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Comment_Pass:Sdcms_Comment_Pass="&t7&vbcrlf
	Config=Config&"'�������"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Comment_IsPass:Sdcms_Comment_IsPass="&t8&vbcrlf
	Config=Config&"'Html��ǩ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadHtml:Sdcms_BadHtml="""&t9&""""&vbcrlf
	Config=Config&"'Html��ǩ�¼�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadEvent:Sdcms_BadEvent="""&t10&""""&vbcrlf
	Config=Config&"'�໰����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_BadText:Sdcms_BadText="""&t11&""""&vbcrlf
	Config=Config&"'��Ϣ�����Զ���ȡ�ַ���,����Ϊ��200-500"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Length:Sdcms_Length="&t22&vbcrlf
	Config=Config&"'���ݿ���������,TrueΪAccess��FalseΪMSSQL"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataType:Sdcms_DataType="&Sdcms_DataType&vbcrlf
	Config=Config&"'Access���ݿ�Ŀ¼"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataFile:Sdcms_DataFile="""&Sdcms_DataFile&""""&vbcrlf
	Config=Config&"'Access���ݿ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_DataName:Sdcms_DataName="""&Sdcms_DataName&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ�IP,������ (local) �����IP "& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlHost:Sdcms_SqlHost="""&Sdcms_SqlHost&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ���"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlData:Sdcms_SqlData="""&Sdcms_SqlData&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ��ʻ�"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlUser:Sdcms_SqlUser="""&Sdcms_SqlUser&""""&vbcrlf
	Config=Config&"'MSSQL���ݿ�����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_SqlPass:Sdcms_SqlPass="""&Sdcms_SqlPass&""""&vbcrlf
	Config=Config&"'ϵͳ��װ����"& vbcrlf
	Config=Config&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_CreateDate:Sdcms_CreateDate="""&Sdcms_CreateDate&""""&vbcrlf
	Config=Config&"%"
	Config=Config&">"&vbcrlf
	Config=Config&"<!--#include file=""../Skins/"&Sdcms_Skins_Root&"/Skins.asp""-->"
	Savefile "../Inc/","Const.asp",Config
	
	'����ˮӡ�����ļ�
	Dim Jpeg
	Jpeg=""
	Jpeg=Jpeg&"<"
	Jpeg=Jpeg&"%"& vbcrlf
	Jpeg=Jpeg&"'ˮӡ����"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t0:Sdcms_Jpeg_t0="&f0&vbcrlf
	Jpeg=Jpeg&"'ˮӡ���ͣ�����Ϊ�棬ͼƬΪ��"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t1:Sdcms_Jpeg_t1="&f1&vbcrlf
	Jpeg=Jpeg&"'ˮӡ����"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t2:Sdcms_Jpeg_t2="""&f2&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ����"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t3:Sdcms_Jpeg_t3="""&f3&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ��������"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t4:Sdcms_Jpeg_t4="""&f4&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ���������С"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t5:Sdcms_Jpeg_t5="""&f5&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ������ɫ"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t6:Sdcms_Jpeg_t6="""&f6&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ�����Ƿ�Ӵ�"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t7:Sdcms_Jpeg_t7="&f7&vbcrlf
	Jpeg=Jpeg&"'ˮӡͼƬ·��"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t8:Sdcms_Jpeg_t8="""&f8&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ͸����"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t9:Sdcms_Jpeg_t9="""&f9&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡͼƬ��ɫ"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t10:Sdcms_Jpeg_t10="""&f10&""""&vbcrlf
	Jpeg=Jpeg&"'ˮӡ�������λ��"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t11:Sdcms_Jpeg_t11="&f11&vbcrlf
	Jpeg=Jpeg&"'ˮӡ����λ��,X Y����|��"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t12:Sdcms_Jpeg_t12=Split("""&f12_0&"|"&f12_1&""",""|"")"&vbcrlf
	Jpeg=Jpeg&"'����ͼ���"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t13:Sdcms_Jpeg_t13="""&IsNum(f13,0)&""""&vbcrlf
	Jpeg=Jpeg&"'����ͼ�߶�"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t14:Sdcms_Jpeg_t14="""&IsNum(f14,0)&""""&vbcrlf
	Jpeg=Jpeg&"'����ͼ�㷨"& vbcrlf
	Jpeg=Jpeg&""&CHR(32)&CHR(32)&CHR(32)&CHR(32)&"Dim Sdcms_Jpeg_t15:Sdcms_Jpeg_t15="""&f15&""""&vbcrlf
	Jpeg=Jpeg&"%"
	Jpeg=Jpeg&">"
	Savefile "../Inc/","AspJpeg.asp",Jpeg
	
	AddLog sdcms_adminname,GetIp,"�޸�ϵͳ����",0
	
	IF Not(t6) Then Application.Contents.RemoveAll()
	
	
	
	IF Change_Mode Then
		IF t23=2 Then
			Alert "ϵͳ���ñ���ɹ���\n\n����ģʽ�ı䣬ϵͳ��ˢ�µ�ǰҳ�档\n\n��������Ҫ�����������ҳ���������ʹ��ǰ̨���ܣ�����","./"
		Else
			Alert "ϵͳ���ñ���ɹ���\n\n����ģʽ�ı䣬ϵͳ��ˢ�µ�ǰҳ�档","./"
		End IF
	Else
		Alert "ϵͳ���ñ���ɹ���","?"
	End IF
End Sub

Function Check_Add
	Check_Add="	<script>"&vbcrlf
	Check_Add=Check_Add&"	function checkadd()"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t0.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��վ���Ʋ���Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t0.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t1.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��վ��������Ϊ��');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t1.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	if (document.add.t3.value=='')"&vbcrlf
	Check_Add=Check_Add&"	{"&vbcrlf
	Check_Add=Check_Add&"	alert('��ѡ���ļ���׺');"&vbcrlf
	Check_Add=Check_Add&"	document.add.t3.focus;"&vbcrlf
	Check_Add=Check_Add&"	return false"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	}"&vbcrlf
	Check_Add=Check_Add&"	</script>"&vbcrlf
End Function
%>  
</div>
</body>
</html>