<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>���������˵�</TITLE>
<META http-equiv=Content-Type content="text/html; charset=GBK">
<LINK href="AdminDefaultTheme/index.css" type=text/css rel=stylesheet>
<LINK href="AdminDefaultTheme/MasterPage.css" type=text/css rel=stylesheet>
<LINK href="AdminDefaultTheme/Guide.css" type=text/css rel=stylesheet>
<script type="text/javascript">
<!--
function Switch(obj)
{
    obj.className = (obj.className == "guideexpand") ? "guidecollapse" : "guideexpand";
    var nextDiv;
    if (obj.nextSibling)
    {
        if(obj.nextSibling.nodeName=="DIV")
        {
            nextDiv = obj.nextSibling;
        }
        else
        {
            if(obj.nextSibling.nextSibling)
            {
                if(obj.nextSibling.nextSibling.nodeName=="DIV")
                {
                    nextDiv = obj.nextSibling.nextSibling;
                }
            }
        }
        if(nextDiv)
        {
            nextDiv.style.display = (nextDiv.style.display != "") ? "" : "none";
        }
    }
}
function Clearhtml()
{
    var bln=confirm("ע�⣺���ӡ��޸ġ�ɾ���������ʱ���Զ����ɡ����¡�ɾ�������ɵľ�̬�ļ���\n�����û�ж�ģ�������޸ģ�����Ҫ��������������Ʒ��������ϸҳ�棡\n��������Բ�Ʒ�����š����ء��˲ŵȷ���ҳ�������޸ģ�ֻ��Ҫ������ط���ҳ�档\n\n��ȷ���Ƿ������");
    return bln;
}
function ClearhtmlAll()
{
    var bln=confirm("���棺��������ȫվ��̬ҳ�潫�ķѽ϶�ϵͳ��Դ��\n��ȷ���Ƿ������");
    return bln;
}
function ChkSlide()
{
    var bln=confirm("ע�⣺�޸���ͼƬ�����󣬱��뷢���õ�Ƭ�Ը���ǰ̨��ʾ��");
    return bln;
}
function ClearCount()
{
    var bln=confirm("���棺�Ƿ�ȷ������û�ͳ�����ݣ�\n��պ󽫲��ָܻ���");
    return bln;
}
-->
</script>
<% ID=Trim(Request.QueryString("ID")) %>
<META content="MSHTML 6.00.6001.18226" name=GENERATOR></HEAD>
<BODY id=Guidebody>
<DIV id=Guide_back>
<UL>
  <LI id=Guide_top>
  <DIV id=Guide_toptext>��ݵ���</DIV>
  <LI id=Guide_main>
  <DIV id=Guide_box>
<% If ID="System" or ID="" Then %>
  <div class="guideexpand" onClick="Switch(this)">ϵͳ��������</div>
  <DIV class=guide>
  <UL id=Links>
    <LI><A href="SetSite.asp" target="main_right">��վ��������</A></li>
    <li><a href="SetConst.Asp" target="main_right">ϵͳ�߼���������</a></li>
    <LI><A href="NavigationEdit.asp?Result=Add" target="main_right">����������</A></li>
    <LI><A href="NavigationList.asp" target="main_right">����������</A></li>
    <LI><A href="FriendLinkEdit.asp?Result=Add" target="main_right">������������</A></li>
    <LI><A href="FriendLinkList.asp" target="main_right">�������ӹ���</A></li>
    <li><a href="SetKey.asp" target="main_right">վ�����ӹ���<font color="red">(New)</font></a></li>
    <li><a href="LinkEdit.asp?Result=Add" target="main_right">վ����������<font color="red">(New)</font></a></li>
    <!--<li><a href="MyEditManage.asp" target="main_right">�ı��༭������<font color="red">(New)</font></a></li>-->
    <li><a href="eWebEditor/Manage/style.Asp" target="main_right" style="color:#ccc" title="�Ѿ�ʧЧ!" onclick="alert('�����Ѿ�ʧЧ!')">�ı��༭������_old</a></li>
    <li><a href="eWebEditor/Manage/upload.Asp" target="main_right">�ϴ�ͼƬ����<font color="red">(New)</font></a></li>
    <li><a href="Admin_SiteMap.asp" target="main_right">���ɹȸ�SiteMap<font color="red">(New)</font></a></li>
    <li><a href="Admin_XML.asp" target="main_right">���ɰٶ�XML<font color="red">(New)</font></a></li>
    <li><a href="UserMessage.Asp" target="main_right">�ͻ���ʱ��ѯ����<font color="red">(New)</font></a></li>
    <li><a href="Admin_Cache.Asp" target="main_right">���ϵͳ����<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Data.Asp?Action=DataBackup" target="main_right">����ϵͳ���ݿ�<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Data.Asp?Action=DataCompact" target="main_right">ѹ�����޸�ϵͳ���ݿ�<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Multilingual" Then %>
        <div class="guideexpand" onClick="Switch(this)">������Թ���</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="ChinaQJ_Multi_Language.Asp" target="main_right">�������ģ�����<font color="red">(New)</font></a></li>
            <li><a href="ChinaQJ_Multi_Language_Edit.Asp?Result=Add" target="main_right">����������ģ��<font color="red">(New)</font></a></li>
            <li><a href="Language.Asp" target="main_right">ϵͳ���԰�����<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="CorporateInformation" Then %>
  <div class="guideexpand" onClick="Switch(this)">��ҵ��Ϣ����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="AboutList.asp" target="main_right">��ҵ��Ϣ�б�</a></li>
    <li><a href="AboutEdit.asp?Result=Add" target="main_right">������ҵ��Ϣ</a></li>
  </ul>
  </div>
<% ElseIf ID="New" Then %>
  <div class="guideexpand" onClick="Switch(this)">������Ѷ����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="NewsSort.asp?Action=Add&ParentID=0" target="main_right">����������</a></li>
    <li><a href="NewsList.asp" target="main_right">�����б�����</a></li>
    <li><a href="NewsEdit.asp?Result=Add" target="main_right">��������</a></li>
    <li><a href="KeyList.Asp" target="main_right">��β�ؼ��ʹ���</a></li>
    <li><a href="KeyEdit.Asp?Result=Add" target="main_right">�������ӳ�β�ؼ���</a></li>
    <li><a href="KeyIdeaList.Asp" target="main_right">��β�ؼ��ʴ������</a></li>
    <li><a href="KeyIdeaEdit.Asp?Result=Add" target="main_right">���ӳ�β�ؼ��ʴ���</a></li>
  </ul>
  </div>
<% ElseIf ID="Product" Then %>
  <div class="guideexpand" onClick="Switch(this)">��˾��Ʒ����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="ProductSort.asp?Action=Add&ParentID=0" target="main_right">��Ʒ������</a></li>
    <li><a href="ProductList.asp" target="main_right">��Ʒ�б�����</a></li>
    <li><a href="ProductEdit.asp?Result=Add" target="main_right">���Ӳ�Ʒ��Ϣ</a></li>
    <li><a href="PropertiesList.Asp" target="main_right">��Ʒ���Թ���</a></li>
    <li><a href="PropertiesEdit.Asp?Result=Add" target="main_right">���Ӳ�Ʒ����</a></li>
  </ul>
  </div>
<% ElseIf ID="Download" Then %>
  <div class="guideexpand" onClick="Switch(this)">�������ع���</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="DownSort.asp?Action=Add&ParentID=0" target="main_right">����������</a></li>
    <li><a href="DownList.asp" target="main_right">�����б�����</a></li>
    <li><a href="DownEdit.asp?Result=Add" target="main_right">����������Ϣ</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Case" Then %>
  <div class="guideexpand" onClick="Switch(this)">�ͻ�����ģ��</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="ImageSort.Asp?Action=Add&ParentID=0" target="main_right">����������</a></li>
            <li><a href="ImageList.Asp" target="main_right">�����б�����</a></li>
            <li><a href="ImageEdit.Asp?Result=Add" target="main_right">���Ӱ�����Ϣ</a></li>
          </ul>
        </div>
<% ElseIf ID="Other" Then %>
  <div class="guideexpand" onClick="Switch(this)">������Ϣ����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="OthersSort.asp?Action=Add&ParentID=0" target="main_right">��Ϣ������</a></li>
    <li><a href="OthersList.asp" target="main_right">��Ϣ�б�����</a></li>
    <li><a href="OthersEdit.asp?Result=Add" target="main_right">������Ϣ</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Talent" Then %>
  <div class="guideexpand" onClick="Switch(this)">�˲���Ƹ����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="JobsList.asp" target="main_right">��Ƹ�б�����</a></li>
    <li><a href="JobsEdit.asp?Result=Add" target="main_right">������Ƹ��Ϣ</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Magazine" Then %>
        <div class="guideexpand" onClick="Switch(this)">������־����</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="MagazineSort.Asp?Action=Add&ParentID=0" target="main_right">������־������</a></li>
            <li><a href="MagazineList.Asp" target="main_right">������־�б�����</a></li>
            <li><a href="MagazineEdit.Asp?Result=Add" target="main_right">���ӵ�����־</a></li>
            <li><a href="MagazineMusic.Asp" target="main_right">������־������������</a></li>
            <li><a href="MagazineSetting.Asp" target="main_right">������־��������</a></li>
          </ul>
        </div>
<% ElseIf ID="Video" Then %>
        <div class="guideexpand" onClick="Switch(this)">��ҵ��Ƶ����</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="VideoSort.Asp?Action=Add&ParentID=0" target="main_right">��ҵ��Ƶ������</a></li>
            <li><a href="VideoList.Asp" target="main_right">��ҵ��Ƶ�б�����</a></li>
            <li><a href="VideoEdit.Asp?Result=Add" target="main_right">������Ƶ��Ϣ</a></li>
          </ul>
        </div>
<% ElseIf ID="Feedback" Then %>
  <div class="guideexpand" onClick="Switch(this)">��ѯ��������</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="MessageList.asp" target="main_right">������Ϣ����</a></li>
    <li><a href="OrderList.asp" target="main_right">������Ϣ����</a></li>
    <li><a href="TalentsList.asp" target="main_right">�˲���Ϣ����</a></li>
	<li><a href="BUG.asp" target="main_right">BUG��������</a></li>
<% ElseIf ID="User" Then %>
  <div class="guideexpand" onClick="Switch(this)">��վ��Ա����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="AdminList.asp" target="main_right">��վ����Ա����</a></li>
    <li><a href="AdminEdit.asp?Result=Add" target="main_right">������վ����Ա</a></li>
    <li><a href="MemList.asp" target="main_right">ǰ̨��Ա����</a></li>
    <li><a href="MemGroup.asp" target="main_right">��Ա������</a></li>
    <li><a href="MemGroup.asp?Result=Add" target="main_right">���ӻ�Ա���</a></li>
    <li><a href="ManageLog.asp" target="main_right">��̨��¼��־����</a></li>
<% ElseIf ID="Html" Then %>
  <div class="guideexpand" onClick="Switch(this)">��̬ҳ�����</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="Admin_Html_Need.Asp" target="main_right" onClick="return Clearhtml()">�������ɾ�̬ҳ��</a></li>
    <li><a href="Admin_html.Asp" target="main_right" onClick="return ClearhtmlAll()"><font color="red">����ȫվ��̬ҳ��</font><font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Plug" Then %>
  <div class="guideexpand" onClick="Switch(this)">�߼���������</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="Admin_Slide.asp" target="main_right">�õ�Ƭ����������<font color="red">(New)</font></a></li>
    <li><a href="Admin_SlideEdit.asp?ShowType=Slide" target="main_right" onClick="return ChkSlide()">Flash�õ�Ƭ����<font color="red">(New)</font></a></li>
    <li><a href="Admin_Search.asp" target="main_right">�û������ؼ���<font color="red">(New)</font></a></li>
    <li><a href="Admin_EMail.asp" target="main_right">�ʼ����Ĺ���<font color="red">(New)</font></a></li>
    <li><a href="Admin_EMailPub.asp" target="main_right">�û��ʼ�Ⱥ��<font color="red">(New)</font></a></li>
    <li><a href="Album.Asp" target="main_right">��ҵFlash���<font color="red">(New)</font></a></li>
    <li><a href="Admin_SubSidiaryList.asp" target="main_right">���ӹ�˾����<font color="red">(New)</font></a></li>
    <li><a href="Admin_SubSidiaryEdit.asp?Result=Add" target="main_right">�����ӹ�˾����<font color="red">(New)</font></a></li>
    <li><a href="Admin_Vote.asp" target="main_right">����ͶƱ����<font color="red">(New)</font></a></li>
    <li><a href="Admin_Vote.asp?Action=Add" target="main_right">���ӵ���ͶƱ<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_NetWorkList.asp" target="main_right">Ӫ���������<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_NetWorkEdit.asp?Result=Add" target="main_right">�����������<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Form_Diy.asp" target="main_right">�Զ����������<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Form_Diy.asp?Action=FormAdd" target="main_right">�����Զ������<font color="red">(New)</font></a></li>
    <li><a href="KefuList.Asp" target="main_right">�������߿ͷ�����<font color="red">(New)</font></a></li>
    <li><a href="KefuEdit.Asp?Result=Add" target="main_right">�����¿ͷ�<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="DiyForm" Then %>
        <div class="guideexpand" onClick="Switch(this)">�Զ�������������</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="UserRegister.Asp" target="main_right">���û�ע���������<font color="red">(New)</font></a></li>
            <li><a href="UserCart.Asp" target="main_right">���ﳵ��������<font color="red">(New)</font></a></li>
            <li><a href="Recruitment.Asp" target="main_right">�˲���Ƹ����<font color="red">(New)</font></a></li>
            <li><a href="MessageForm.Asp" target="main_right">�û����ԡ���ѯ����<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Count" Then %>
  <div class="guideexpand" onClick="Switch(this)">����ͳ�ƹ���</div>
  <DIV class=guide>
  <ul id="Links">
    <li><a href="Admin_Count.asp" target="main_right">ͳ�Ƹſ�<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=all" target="main_right">��ϸͳ������<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=chour" target="main_right">���24Сʱͳ��<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cday" target="main_right">����ͳ������<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cweek" target="main_right">��ͳ������<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cmonth" target="main_right">��ͳ������<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=ccome" target="main_right">�û���Դͳ��<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cpage" target="main_right">�û�����ҳ��<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cip" target="main_right">��������ͳ��<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=del" target="main_right" onClick="return ClearCount()">�������ͳ������<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine" Then %>
  <div class="guideexpand" onClick="Switch(this)">���������¼</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.baidu.com/search/url_submit.html" target="_blank">�ٶȵ�¼���</a></li>
            <li><a href="http://www.google.com/intl/zh-CN/add_url.html" target="_blank">Google��¼���</a></li>
            <li><a href="http://search.help.cn.yahoo.com/h4_4.html" target="_blank">Yahoo��¼���</a></li>
            <li><a href="http://search.msn.com/docs/submit.Aspx" target="_blank">Live��¼���</a></li>
            <li><a href="http://www.dmoz.org/World/Chinese_Simplified/" target="_blank">Dmoz��¼���</a></li>
            <li><a href="http://www.alexa.com/site/help/webmasters" target="_blank">Alexa��¼���</a></li>
            <li><a href="http://ads.zhongsou.com/register/page.jsp" target="_blank">���ѵ�¼���</a></li>
            <li><a href="http://iask.com/guest/add_url.php" target="_blank">���ʵ�¼���</a></li>
            <li><a href="http://tellbot.youdao.com/report" target="_blank">�е���¼���</a></li>
            <li><a href="http://cn.bing.com/docs/submit.Aspx" target="_blank">��Ӧ��¼���</a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine2" Then %>
        <div class="guideexpand" onClick="Switch(this)">��ҵ��Ϣ����</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.alibaba.com" target="_blank">����Ͱ�</a></li>
            <li><a href="http://www.hc360.com" target="_blank">Google�۴���</a></li>
            <li><a href="http://www.yp.net.cn" target="_blank">�й���ҳ����</a></li>
            <li><a href="http://yp.sina.net" target="_blank">������ҵ��ҳ</a></li>
            <li><a href="http://www.made-in-china.com" target="_blank">Made-in-China</a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine3" Then %>
        <div class="guideexpand" onClick="Switch(this)">������Ϣ����</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.58.com" target="_blank">58ͬ�Ƿ���</a></li>
            <li><a href="http://www.koubei.com" target="_blank">Yahoo�ڱ�</a></li>
            <li><a href="http://www.ganji.com" target="_blank">�ϼ���</a></li>
            <li><a href="http://www.bendibao.com" target="_blank">���ر�</a></li>
            <li><a href="http://www.baixin.com" target="_blank">������</a></li>
            <li><a href="http://www.fenlei168.com" target="_blank">�й�������Ϣ��</a></li>
          </ul>
        </div>
<% Else %>
<% End If %>
  </div></li></li>
  <LI id=Guide_bottom></LI>
</BODY></HTML>