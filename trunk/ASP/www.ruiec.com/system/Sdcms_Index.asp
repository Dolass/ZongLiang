<!--#include file="sdcms_check.asp"-->
<!--#include file="../plug/Config.asp"-->
<%
Dim Sdcms
Set Sdcms=New Sdcms_Admin
Sdcms.Check_admin
Sdcms.Check_lever ""
Set Sdcms=Nothing
Sdcms_Head
%>
<script type="text/javascript">
	$(document).ready(
	function() 
	{
		$(".left_title").click(function(){
			$(this).next("ul").slideToggle()
			.siblings(".dis:visible").slideUp();
			$(this).toggleClass("left_title_over");
			$(this).siblings(".left_title_over").removeClass("left_title_over");
		});
	});
</script>
<Script>if(self!=top){top.location=self.location;}</script>
	<div id="head">
		<div class="left"><img src="images/sdcms_logo.gif" alt="����ý��վ��Ϣ����ϵͳ"/></div>
		<div class="left head_txt">���ã�<%=sdcms_adminname%>��[ <a href="sdcms_admin.asp?id=<%=sdcms_adminid%>&action=edit" target="main">�ҵ��ʻ�</a> <span><a href="index.asp?action=out">�˳�</a></span> ]</div>
		<div class="right head_menu">
		  <ul id="head_menu">
			  <li><a href="../" target="_blank">Ԥ����վ</a></li>
			  <li><a href="sdcms_index.asp">ˢ�º�̨</a></li>
			  <li><a href="sdcms_set.asp" target="main">ϵͳ����</a></li>
			  <li><a href="sdcms_info.asp" target="main">��Ϣ����</a></li>
			  <li><a href="sdcms_cache.asp" target="main">���»���</a></li>
		  </ul>
		</div>	
	</div>
<!--head is over-->
	 <div id="content">
	 <div id="left">
          <div class="left_title">ϵͳ����</div>
		  <ul class="dis">
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_set.asp" target="main">ϵͳ����</a>����<a href="sdcms_log.asp" target="main">��־</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_admin.asp?action=add" target="main">����ʻ�</a>����<a href="sdcms_admin.asp" target="main">����</a></li>
		  </ul>
		  
		  <div class="left_title">��Ϣ����</div>
		  <ul class="dis">
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_class.asp?action=add" target="main">��ӷ���</a>����<a href="sdcms_class.asp" target="main">����</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="Sdcms_Topic.asp?action=add" target="main">���ר��</a>����<a href="Sdcms_Topic.asp" target="main">����</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_info.asp?action=add" target="main">�����Ϣ</a>����<a href="sdcms_info.asp" target="main">����</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_Page.asp?action=add" target="main">��ӵ�ҳ</a>����<a href="sdcms_Page.asp" target="main">����</a></li>
		  </ul>
		  
		  <div class="left_title">���ӹ���</div>
		  <ul class="dis">
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_sitelink.asp" target="main">��������</a>��<a href="sdcms_tags.asp" target="main">Tags����</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_search.asp" target="main">��������</a>��<a href="sdcms_outsite.asp" target="main">վ�����</a></li>
		  </ul>
		  
		  <div class="left_title">�������</div>
		  <ul class="dis">
		  <%=Plug_menu%>
		  </ul>
		  
		  <div class="left_title">�������</div>
		  <ul class="dis">
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_label.asp?action=add" target="main">�����Ƭ</a>����<a href="sdcms_label.asp" target="main">����</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_skins.asp" target="main">��վģ�����</a></li>
		  </ul>
		  <%IF Sdcms_Mode=2 Then%>
		  <div class="left_title">���ɹ���</div>
		  <ul class="dis">
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_create.asp?Stype=1"  target="main">������ҳ</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_create.asp?Stype=2"  target="main">������Ŀ</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_create.asp?Stype=3"  target="main">������Ϣ</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_create.asp?Stype=4"  target="main">���ɵ�ҳ</a></li>
		  <li class="left_link" onClick="DoLocation(this)"><a href="sdcms_create.asp?Stype=5"  target="main">���ɵ�ͼ</a></li>
		  </ul>  
		  <%End IF%>
     </div>
	   
	<div id="right">
		<iframe id="Main_Content" scrolling="auto" name="main" src="sdcms_main.asp" frameborder="0"></iframe>
	</div>

	</div>
</div>
<script language="javascript">window.setInterval("reinitIframe('Main_Content')",300);</script>
</body>
</html>
