<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>管理导航菜单</TITLE>
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
    var bln=confirm("注意：添加、修改、删除相关数据时会自动生成、更新、删除所生成的静态文件。\n如果您没有对模板作过修改，不需要批量生成所有商品或新闻详细页面！\n如果您仅对产品、新闻、下载、人才等分类页面作过修改，只需要生成相关分类页面。\n\n请确定是否操作？");
    return bln;
}
function ClearhtmlAll()
{
    var bln=confirm("警告：批量生成全站静态页面将耗费较多系统资源！\n请确定是否操作？");
    return bln;
}
function ChkSlide()
{
    var bln=confirm("注意：修改完图片参数后，必须发布幻灯片以更新前台显示！");
    return bln;
}
function ClearCount()
{
    var bln=confirm("警告：是否确定清空用户统计数据！\n清空后将不能恢复！");
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
  <DIV id=Guide_toptext>快捷导航</DIV>
  <LI id=Guide_main>
  <DIV id=Guide_box>
<% If ID="System" or ID="" Then %>
  <div class="guideexpand" onClick="Switch(this)">系统参数设置</div>
  <DIV class=guide>
  <UL id=Links>
    <LI><A href="SetSite.asp" target="main_right">网站参数设置</A></li>
    <li><a href="SetConst.Asp" target="main_right">系统高级参数设置</a></li>
    <LI><A href="NavigationEdit.asp?Result=Add" target="main_right">导航栏添加</A></li>
    <LI><A href="NavigationList.asp" target="main_right">导航栏管理</A></li>
    <LI><A href="FriendLinkEdit.asp?Result=Add" target="main_right">友情链接添加</A></li>
    <LI><A href="FriendLinkList.asp" target="main_right">友情链接管理</A></li>
    <li><a href="SetKey.asp" target="main_right">站内链接管理<font color="red">(New)</font></a></li>
    <li><a href="LinkEdit.asp?Result=Add" target="main_right">站内链接添加<font color="red">(New)</font></a></li>
    <!--<li><a href="MyEditManage.asp" target="main_right">文本编辑器管理<font color="red">(New)</font></a></li>-->
    <li><a href="eWebEditor/Manage/style.Asp" target="main_right" style="color:#ccc" title="已经失效!" onclick="alert('设置已经失效!')">文本编辑器管理_old</a></li>
    <li><a href="eWebEditor/Manage/upload.Asp" target="main_right">上传图片管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_SiteMap.asp" target="main_right">生成谷歌SiteMap<font color="red">(New)</font></a></li>
    <li><a href="Admin_XML.asp" target="main_right">生成百度XML<font color="red">(New)</font></a></li>
    <li><a href="UserMessage.Asp" target="main_right">客户即时咨询管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_Cache.Asp" target="main_right">清除系统缓存<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Data.Asp?Action=DataBackup" target="main_right">备份系统数据库<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Data.Asp?Action=DataCompact" target="main_right">压缩、修复系统数据库<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Multilingual" Then %>
        <div class="guideexpand" onClick="Switch(this)">多国语言管理</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="ChinaQJ_Multi_Language.Asp" target="main_right">多国语言模块管理<font color="red">(New)</font></a></li>
            <li><a href="ChinaQJ_Multi_Language_Edit.Asp?Result=Add" target="main_right">添加新语言模块<font color="red">(New)</font></a></li>
            <li><a href="Language.Asp" target="main_right">系统语言包管理<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="CorporateInformation" Then %>
  <div class="guideexpand" onClick="Switch(this)">企业信息管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="AboutList.asp" target="main_right">企业信息列表</a></li>
    <li><a href="AboutEdit.asp?Result=Add" target="main_right">添加企业信息</a></li>
  </ul>
  </div>
<% ElseIf ID="New" Then %>
  <div class="guideexpand" onClick="Switch(this)">新闻资讯管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="NewsSort.asp?Action=Add&ParentID=0" target="main_right">新闻类别管理</a></li>
    <li><a href="NewsList.asp" target="main_right">新闻列表管理</a></li>
    <li><a href="NewsEdit.asp?Result=Add" target="main_right">添加新闻</a></li>
    <li><a href="KeyList.Asp" target="main_right">长尾关键词管理</a></li>
    <li><a href="KeyEdit.Asp?Result=Add" target="main_right">批量添加长尾关键词</a></li>
    <li><a href="KeyIdeaList.Asp" target="main_right">长尾关键词创意管理</a></li>
    <li><a href="KeyIdeaEdit.Asp?Result=Add" target="main_right">添加长尾关键词创意</a></li>
  </ul>
  </div>
<% ElseIf ID="Product" Then %>
  <div class="guideexpand" onClick="Switch(this)">公司产品管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="ProductSort.asp?Action=Add&ParentID=0" target="main_right">产品类别管理</a></li>
    <li><a href="ProductList.asp" target="main_right">产品列表管理</a></li>
    <li><a href="ProductEdit.asp?Result=Add" target="main_right">添加产品信息</a></li>
    <li><a href="PropertiesList.Asp" target="main_right">产品属性管理</a></li>
    <li><a href="PropertiesEdit.Asp?Result=Add" target="main_right">添加产品属性</a></li>
  </ul>
  </div>
<% ElseIf ID="Download" Then %>
  <div class="guideexpand" onClick="Switch(this)">资料下载管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="DownSort.asp?Action=Add&ParentID=0" target="main_right">下载类别管理</a></li>
    <li><a href="DownList.asp" target="main_right">下载列表管理</a></li>
    <li><a href="DownEdit.asp?Result=Add" target="main_right">添加下载信息</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Case" Then %>
  <div class="guideexpand" onClick="Switch(this)">客户案例模块</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="ImageSort.Asp?Action=Add&ParentID=0" target="main_right">案例类别管理</a></li>
            <li><a href="ImageList.Asp" target="main_right">案例列表管理</a></li>
            <li><a href="ImageEdit.Asp?Result=Add" target="main_right">添加案例信息</a></li>
          </ul>
        </div>
<% ElseIf ID="Other" Then %>
  <div class="guideexpand" onClick="Switch(this)">其他信息管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="OthersSort.asp?Action=Add&ParentID=0" target="main_right">信息类别管理</a></li>
    <li><a href="OthersList.asp" target="main_right">信息列表管理</a></li>
    <li><a href="OthersEdit.asp?Result=Add" target="main_right">添加信息</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Talent" Then %>
  <div class="guideexpand" onClick="Switch(this)">人才招聘管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="JobsList.asp" target="main_right">招聘列表管理</a></li>
    <li><a href="JobsEdit.asp?Result=Add" target="main_right">添加招聘信息</a></li>
    </UL>
    </DIV>
<% ElseIf ID="Magazine" Then %>
        <div class="guideexpand" onClick="Switch(this)">电子杂志管理</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="MagazineSort.Asp?Action=Add&ParentID=0" target="main_right">电子杂志类别管理</a></li>
            <li><a href="MagazineList.Asp" target="main_right">电子杂志列表管理</a></li>
            <li><a href="MagazineEdit.Asp?Result=Add" target="main_right">添加电子杂志</a></li>
            <li><a href="MagazineMusic.Asp" target="main_right">电子杂志背景音乐设置</a></li>
            <li><a href="MagazineSetting.Asp" target="main_right">电子杂志参数设置</a></li>
          </ul>
        </div>
<% ElseIf ID="Video" Then %>
        <div class="guideexpand" onClick="Switch(this)">企业视频管理</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="VideoSort.Asp?Action=Add&ParentID=0" target="main_right">企业视频类别管理</a></li>
            <li><a href="VideoList.Asp" target="main_right">企业视频列表管理</a></li>
            <li><a href="VideoEdit.Asp?Result=Add" target="main_right">添加视频信息</a></li>
          </ul>
        </div>
<% ElseIf ID="Feedback" Then %>
  <div class="guideexpand" onClick="Switch(this)">咨询反馈管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="MessageList.asp" target="main_right">留言信息管理</a></li>
    <li><a href="OrderList.asp" target="main_right">订单信息管理</a></li>
    <li><a href="TalentsList.asp" target="main_right">人才信息管理</a></li>
<% ElseIf ID="User" Then %>
  <div class="guideexpand" onClick="Switch(this)">网站会员管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="AdminList.asp" target="main_right">网站管理员管理</a></li>
    <li><a href="AdminEdit.asp?Result=Add" target="main_right">添加网站管理员</a></li>
    <li><a href="MemList.asp" target="main_right">前台会员资料</a></li>
    <li><a href="MemGroup.asp" target="main_right">会员组别管理</a></li>
    <li><a href="MemGroup.asp?Result=Add" target="main_right">添加会员组别</a></li>
    <li><a href="ManageLog.asp" target="main_right">后台登录日志管理</a></li>
<% ElseIf ID="Html" Then %>
  <div class="guideexpand" onClick="Switch(this)">静态页面管理</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="Admin_Html_Need.Asp" target="main_right" onClick="return Clearhtml()">按需生成静态页面</a></li>
    <li><a href="Admin_html.Asp" target="main_right" onClick="return ClearhtmlAll()"><font color="red">生成全站静态页面</font><font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Plug" Then %>
  <div class="guideexpand" onClick="Switch(this)">高级功能设置</div>
  <DIV class=guide>
  <UL id=Links>
    <li><a href="Admin_Slide.asp" target="main_right">幻灯片参数及发布<font color="red">(New)</font></a></li>
    <li><a href="Admin_SlideEdit.asp?ShowType=Slide" target="main_right" onClick="return ChkSlide()">Flash幻灯片管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_Search.asp" target="main_right">用户搜索关键词<font color="red">(New)</font></a></li>
    <li><a href="Admin_EMail.asp" target="main_right">邮件订阅管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_EMailPub.asp" target="main_right">用户邮件群发<font color="red">(New)</font></a></li>
    <li><a href="Album.Asp" target="main_right">企业Flash相册<font color="red">(New)</font></a></li>
    <li><a href="Admin_SubSidiaryList.asp" target="main_right">多子公司管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_SubSidiaryEdit.asp?Result=Add" target="main_right">添加子公司资料<font color="red">(New)</font></a></li>
    <li><a href="Admin_Vote.asp" target="main_right">调查投票管理<font color="red">(New)</font></a></li>
    <li><a href="Admin_Vote.asp?Action=Add" target="main_right">添加调查投票<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_NetWorkList.asp" target="main_right">营销网络管理<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_NetWorkEdit.asp?Result=Add" target="main_right">添加网络管理<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Form_Diy.asp" target="main_right">自定义表单管理<font color="red">(New)</font></a></li>
    <li><a href="ChinaQJ_Form_Diy.asp?Action=FormAdd" target="main_right">添加自定义表单<font color="red">(New)</font></a></li>
    <li><a href="KefuList.Asp" target="main_right">悬浮在线客服管理<font color="red">(New)</font></a></li>
    <li><a href="KefuEdit.Asp?Result=Add" target="main_right">添加新客服<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="DiyForm" Then %>
        <div class="guideexpand" onClick="Switch(this)">自定义表单必填参数</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="UserRegister.Asp" target="main_right">新用户注册表单参数<font color="red">(New)</font></a></li>
            <li><a href="UserCart.Asp" target="main_right">购物车表单参数<font color="red">(New)</font></a></li>
            <li><a href="Recruitment.Asp" target="main_right">人才招聘参数<font color="red">(New)</font></a></li>
            <li><a href="MessageForm.Asp" target="main_right">用户留言、咨询参数<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="Count" Then %>
  <div class="guideexpand" onClick="Switch(this)">流量统计管理</div>
  <DIV class=guide>
  <ul id="Links">
    <li><a href="Admin_Count.asp" target="main_right">统计概况<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=all" target="main_right">详细统计数据<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=chour" target="main_right">最近24小时统计<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cday" target="main_right">今日统计数据<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cweek" target="main_right">周统计数据<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cmonth" target="main_right">月统计数据<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=ccome" target="main_right">用户来源统计<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cpage" target="main_right">用户访问页面<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=cip" target="main_right">来自区域统计<font color="red">(New)</font></a></li>
    <li><a href="Admin_Count.asp?Action=del" target="main_right" onClick="return ClearCount()">清空所有统计数据<font color="red">(New)</font></a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine" Then %>
  <div class="guideexpand" onClick="Switch(this)">搜索引擎登录</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.baidu.com/search/url_submit.html" target="_blank">百度登录入口</a></li>
            <li><a href="http://www.google.com/intl/zh-CN/add_url.html" target="_blank">Google登录入口</a></li>
            <li><a href="http://search.help.cn.yahoo.com/h4_4.html" target="_blank">Yahoo登录入口</a></li>
            <li><a href="http://search.msn.com/docs/submit.Aspx" target="_blank">Live登录入口</a></li>
            <li><a href="http://www.dmoz.org/World/Chinese_Simplified/" target="_blank">Dmoz登录入口</a></li>
            <li><a href="http://www.alexa.com/site/help/webmasters" target="_blank">Alexa登录入口</a></li>
            <li><a href="http://ads.zhongsou.com/register/page.jsp" target="_blank">中搜登录入口</a></li>
            <li><a href="http://iask.com/guest/add_url.php" target="_blank">爱问登录入口</a></li>
            <li><a href="http://tellbot.youdao.com/report" target="_blank">有道登录入口</a></li>
            <li><a href="http://cn.bing.com/docs/submit.Aspx" target="_blank">必应登录入口</a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine2" Then %>
        <div class="guideexpand" onClick="Switch(this)">企业信息发布</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.alibaba.com" target="_blank">阿里巴巴</a></li>
            <li><a href="http://www.hc360.com" target="_blank">Google慧聪网</a></li>
            <li><a href="http://www.yp.net.cn" target="_blank">中国黄页在线</a></li>
            <li><a href="http://yp.sina.net" target="_blank">新浪企业黄页</a></li>
            <li><a href="http://www.made-in-china.com" target="_blank">Made-in-China</a></li>
          </ul>
        </div>
<% ElseIf ID="SearchEngine3" Then %>
        <div class="guideexpand" onClick="Switch(this)">分类信息发布</div>
        <div class="guide">
          <ul id="Links">
            <li><a href="http://www.58.com" target="_blank">58同城分类</a></li>
            <li><a href="http://www.koubei.com" target="_blank">Yahoo口碑</a></li>
            <li><a href="http://www.ganji.com" target="_blank">赶集网</a></li>
            <li><a href="http://www.bendibao.com" target="_blank">本地宝</a></li>
            <li><a href="http://www.baixin.com" target="_blank">百姓网</a></li>
            <li><a href="http://www.fenlei168.com" target="_blank">中国分类信息网</a></li>
          </ul>
        </div>
<% Else %>
<% End If %>
  </div></li></li>
  <LI id=Guide_bottom></LI>
</BODY></HTML>
