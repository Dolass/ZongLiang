﻿<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Version.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>
<script language="javascript" src="JavaScript/Tab.js"></script>
<script language="javascript" src="/Scripts/MyEditObj.js?type=<%=editType%>"></script>
</head>
<%
	
	'Dim obj=Request.Form
	
	'Response.end

%>
<script language="javascript">
<!--
function SiteLogo(){
    var arr = showModalDialog("eWebEditor/customDialog/img.htm", "", "dialogWidth:30em; dialogHeight:26em; status:0;help=no");
    if (arr ==null){
        alert("系统提示：当前没有上传图片，界面预览图为空，用户可以重新上传图片！");
    }
    if (arr !=null){
        editForm.SiteLogo.value=arr;
    }
}
//-->
</script>
<script language="javascript">
<!--
function showUploadDialog(s_Type, s_Link, s_Thumbnail){
    var arr = showModalDialog("eWebEditor/dialog/i_upload.htm?style=coolblue&type="+s_Type+"&link="+s_Link+"&thumbnail="+s_Thumbnail, window, "dialogWidth: 0px; dialogHeight: 0px; help: no; scroll: no; status: no");
}
function changedbtype(dbtype){
  var accesstr = document.getElementById("accesstr");
  var sqltr = document.getElementById("sqltr");
  if(dbtype == 0){
    accesstr.style.display = '';
    sqltr.style.display = 'none';
  }else{
    accesstr.style.display = 'none';
    sqltr.style.display = '';
  }
}
function get(){
    document.getElementById("ProInfo").value=document.getElementById("ProInfoline").value*document.getElementById("ProInfoColumn").value;
}
//-->
</script>
<%
if Instr(session("AdminPurview"),"|1,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end If
select case request.QueryString("Action")
  case "Save"
    SaveSiteInfo
  case "SaveConst"
    SaveConstInfo
  case else
    ViewSiteInfo
end select
%>
<ul id="tablist">
<% 
set rs = server.createobject("adodb.recordset")
sql="select * from ChinaQJ_Language order by ChinaQJ_Language_Order"
rs.open sql,conn,1,1
lani=1
while(not rs.eof)
%>
   <li <% If rs("ChinaQJ_Language_Index") Then Response.Write("class=""selected""")%>><a rel="tcontent<%= lani %>" style="cursor: hand;"><%=rs("ChinaQJ_Language_Name")%></a></li>   
<% 
rs.movenext
lani=lani+1
wend
rs.close
set rs=nothing
%>
</ul>
<body onLoad="changedbtype(0);">
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="0" cellspacing="1">
  <form name="editForm" method="post" action="SetSite.asp?Action=Save">
    <tr>
      <th height="22" colspan="2" sytle="line-height:150%">【系统主参数设置】</th>
    </tr>
    <tr height="28">
      <td colspan="2">
<% 
set rs = server.createobject("adodb.recordset")
sql="select * from ChinaQJ_Language order by ChinaQJ_Language_Order"
rs.open sql,conn,1,1
lani=1
while(not rs.eof)
%>
<div id="tcontent<%= lani %>" class="tabcontent">
<table class="tableborderOther" width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr height="35">
      <td width="200" align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>网站标题：</td>
      <td width="80%" class="forumRowHighlightOther"><input name="SiteTitle<%= rs("ChinaQJ_Language_File") %>" type="text" id="SiteTitle<%= rs("ChinaQJ_Language_File") %>" style="width: 280" value="<%=eval("SiteTitle"&rs("ChinaQJ_Language_File")) %>">        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>关键字：<br />(Keywords)</td>
      <td class="forumRowHighlightOther"><textarea name="Keywords<%= rs("ChinaQJ_Language_File") %>" rows="6"  id="Keywords<%= rs("ChinaQJ_Language_File") %>" style="width: 500"><%=eval("Keywords" & rs("ChinaQJ_Language_File")) %></textarea>        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>网站描述：<br />(Descriptions)</td>
      <td class="forumRowHighlightOther"><textarea name="Descriptions<%= rs("ChinaQJ_Language_File") %>" rows="6" id="Descriptions<%= rs("ChinaQJ_Language_File") %>" style="width: 500"><%=eval("Descriptions" & rs("ChinaQJ_Language_File")) %></textarea>        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>公司名称：</td>
      <td class="forumRowHighlightOther"><input name="ComName<%= rs("ChinaQJ_Language_File") %>" type="text" id="ComName<%= rs("ChinaQJ_Language_File") %>" style="width: 280" value="<%=eval("ComName" & rs("ChinaQJ_Language_File")) %>">        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>公司地址：</td>
      <td class="forumRowHighlightOther"><input name="Address<%= rs("ChinaQJ_Language_File") %>" type="text" id="Address<%= rs("ChinaQJ_Language_File") %>" style="width: 280" value="<%=eval("Address" & rs("ChinaQJ_Language_File")) %>">        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>首页公告图片：</td>
      <td class="forumRowHighlightOther"><input name="SiteIndexNotice<%= rs("ChinaQJ_Language_File") %>" type="text" id="SiteIndexNotice<%= rs("ChinaQJ_Language_File") %>" style="width: 350" value="<%= eval("SiteIndexNotice" & rs("ChinaQJ_Language_File")) %>">
        <input type="button" value="上传图片" onClick="showUploadDialog('image', 'editForm.SiteIndexNotice<%= rs("ChinaQJ_Language_File") %>', '')">        <font color="red">* 默认大小522*283像素</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>首页视频地址：</td>
      <td class="forumRowHighlightOther"><textarea name="Video<%= rs("ChinaQJ_Language_File") %>" rows="6"  id="Video<%= rs("ChinaQJ_Language_File") %>" style="width: 500px;"><%=eval("Video" & rs("ChinaQJ_Language_File")) %></textarea>
        <input type="button" value="上传视频" onClick="showUploadDialog('media', 'editForm.Video<%= rs("ChinaQJ_Language_File") %>', '')">
        <input type="button" value="上传动画" onClick="showUploadDialog('flash', 'editForm.Video<%= rs("ChinaQJ_Language_File") %>', '')">
        <br />
        <font color="#CC0000">视频格式支持：mp3/avi/wmv/asf/mov/rm/ra/ram/rmvb/swf</font> 中文</td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>内页Flash/图片广告设置：</td>
      <td class="forumRowHighlightOther"><input name="PageBanner<%= rs("ChinaQJ_Language_File") %>" type="text" id="PageBanner<%= rs("ChinaQJ_Language_File") %>" style="width: 350" value="<%= eval("PageBanner" & rs("ChinaQJ_Language_File")) %>">
        <input type="button" value="上传图片" onClick="showUploadDialog('image', 'editForm.PageBanner<%= rs("ChinaQJ_Language_File") %>', '')">
        <input type="button" value="上传动画" onClick="showUploadDialog('flash', 'editForm.PageBanner<%= rs("ChinaQJ_Language_File") %>', '')">
        <br />
        <input name="PageBannerType<%= rs("ChinaQJ_Language_File") %>" type="radio" value="0" <% If eval("PageBannerType"&rs("ChinaQJ_Language_File"))=0 Then Response.Write("checked")%>> 显示为Flash动画
        <input name="PageBannerType<%= rs("ChinaQJ_Language_File") %>" type="radio" value="1" <% If eval("PageBannerType"&rs("ChinaQJ_Language_File")) Then Response.Write("checked")%>> 显示为图片广告
        广告宽度：<input name="PageBannerWidth<%= rs("ChinaQJ_Language_File") %>" type="text" id="PageBannerWidth<%= rs("ChinaQJ_Language_File") %>" style="width: 50" value="<%= eval("PageBannerWidth" & rs("ChinaQJ_Language_File")) %>"> 像素
        广告高度：<input name="PageBannerHeight<%= rs("ChinaQJ_Language_File") %>" type="text" id="PageBannerHeight<%= rs("ChinaQJ_Language_File") %>" style="width: 50" value="<%= eval("PageBannerHeight" & rs("ChinaQJ_Language_File")) %>"> 像素 <font color="red">*</font> <br />
        <font color="#CC0000">请选择内页显示的Banner类别，默认为Flash动画。</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRowOther"><%=rs("ChinaQJ_Language_Name")%>首页公司简介：</td>
      <td class="forumRowHighlightOther">
		<div id="div_Con_<%=rs("ChinaQJ_Language_File")%>" style="display:none;"><%= eval("SiteDetail" & rs("ChinaQJ_Language_File")) %></div>
		<script>Start_MyEdit("SiteDetail<%=rs("ChinaQJ_Language_File")%>","div_Con_<%=rs("ChinaQJ_Language_File")%>");</script>
	 </td>
    </tr>
    </table>
    </div>
<% 
rs.movenext
lani=lani+1
wend
rs.close
set rs=nothing
%>
    <script>showtabcontent("tablist")</script></td></tr>
    <tr height="35">
      <td width="200" align="right" class="forumRow">公司网址：</td>
      <td class="forumRowHighlight"><input name="SiteUrl" type="text" id="SiteUrl" style="width: 280" value="<%=SiteUrl%>">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">邮政编码：</td>
      <td class="forumRowHighlight"><input name="ZipCode" type="text" id="ZipCode" style="width: 180" value="<%=ZipCode%>" maxlength="6">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">联系电话：</td>
      <td class="forumRowHighlight"><input name="Telephone" type="text" id="Telephone" style="width: 180" value="<%=Telephone%>">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">传真号码：</td>
      <td class="forumRowHighlight"><input name="Fax" type="text" id="Fax" style="width: 180" value="<%=Fax%>">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">联系手机：</td>
      <td class="forumRowHighlight"><input name="Telephone2" type="text" id="Telephone2" style="width: 180" value="<%=Telephone2%>">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">电子邮箱：</td>
      <td class="forumRowHighlight"><input name="Email" type="text" id="Email" style="width: 180" value="<%=Email%>">
        <font color="red">*</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">ICP备案号：</td>
      <td class="forumRowHighlight"><input name="IcpNumber" type="text" id="IcpNumber" style="width: 180" value="<%=IcpNumber%>"></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">客户留言设置：</td>
      <td class="forumRowHighlight"><input name="MesViewFlag" type="checkbox" id="MesViewFlag" value="1" <%if MesViewFlag then response.write ("checked")%>>
        无须审核</td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">首页Logo设置：</td>
      <td class="forumRowHighlight"><input name="SiteLogo" type="text" id="SiteLogo" style="width: 350" value="<%= SiteLogo %>">
      <input type="button" value="上传图片" onClick="showUploadDialog('image', 'editForm.SiteLogo', '')"></td>
    </tr>
    <tr height="28">
      <td align="right" class="forumRow">企业地图坐标设置：</td>
      <td class="forumRowHighlight">经度：<input name="Longitude" type="text" id="Longitude" style="width: 120px;" value="<%=Longitude%>"> 
        纬度：
          <input name="Latitude" type="text" id="Latitude" style="width: 120px;" value="<%=Latitude%>"> 视图比例：<input name="Proportion" type="text" id="Proportion" style="width: 50px;" value="<%=Proportion%>">
        <a href="javascript: OpenScript('Google_Map_Api.Asp?Longitude=<%=Longitude%>&Latitude=<%=Latitude%>&Proportion=<%=Proportion%>&GoogleMapKey=<%=GoogleMapKey%>',640,480,'')"><img src="Images/icon_qi.gif" align="absmiddle" border="0">修改地图位置</a> <a href="../uploadfile/DemoMap.png" rel="Qianbo[title=Google三维企业地图效果展示]"><u>【Google企业地图效果】</u></a><br />
        Google Maps API 密钥：<input name="GoogleMapKey" type="text" id="GoogleMapKey" style="width: 500px;" value="<%=GoogleMapKey%>">
        <font color="red">*<br />
        Google 地图是一项网络地图服务。通过使用 Google 地图，您可以查询详细地址、寻找周边信息、商户信息，并规划点到点路线。<br />
        您可以方便地利用此项功能进行您公司具体位置的标注，前台将会实时根据地图、卫星、地形显示您的详细地址数据。<br />
        <a href="http://code.google.com/intl/zh-CN/apis/maps/signup.html" target="_blank"><u>注册使用 Google 地图 API(获取Google Maps API 密钥)</u></a>
      </font></td>
    </tr>
    <tr height="28">
      <td align="right" class="forumRow">企业地图默认显示类型：</td>
      <td class="forumRowHighlight"><select style="width: 350px;" name="GoogleMapType">
          <option value="0" <% if GoogleMapType="0" then Response.Write("selected") %> >默认视图</option>
          <option value="1" <% if GoogleMapType="1" then Response.Write("selected") %> >显示 Google 地球卫星图像</option>
          <option value="2" <% if GoogleMapType="2" then Response.Write("selected") %> >混合显示普通视图和卫星视图</option>
        </select>
        <font color="red">*
      <br />每个地图还包含一个“缩放级别”，用于定义当前视图的分辨率。在普通地图视图内，可以使用 0（最低缩放级别，在地图上可以看到整个世界）到 19（最高缩放级别，可以看到独立建筑物）之间的缩放级别。缩放级别因所查看地区而异，因为地球上某些地区的数据比其他地区更详细。在卫星视图中可以使用多达 20 个缩放级别。</font></td>
    </tr>
    <tr height="28">
      <td align="right" class="forumRow">企业地图企业形象图文说明：</td>
      <td class="forumRowHighlight"><textarea name="GoogleMapHtml" rows="6"  id="GoogleMapHtml" style="width: 500px;"><%=GoogleMapHtml%></textarea>
        <font color="red">*
      <br />支持Html代码，或在企业简介编辑器中可视化编辑，将编辑完成的代码复制到本接口。</font></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">统计代码调用：</td>
      <td class="forumRowHighlight"><textarea name="ChinaQJ_Stat" rows="6" id="ChinaQJ_Stat" style="width: 500"><%= ChinaQJ_Stat %></textarea><br />
      <font color="#CC0000">将您申请的第三方统计代码复制到本栏</font>(为您提供网站的流量变化趋势，以及各种重要的流量数据。)<br />
      推荐申请：<a href="http://www.cnzz.com/" target="_blank"><u>站长统计</u></a> <a href="http://www.51.la/" target="_blank"><u>我要啦免费统计</u></a> <a href="http://tongji.cn.yahoo.com/" target="_blank"><u>雅虎统计</u></a> <a href="http://count.51yes.com/" target="_blank"><u>51Yes网站流量统计</u></a></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow">公司简介混合图片：</td>
      <td class="forumRowHighlight"><input name="SiteDetailPic" type="text" id="SiteDetailPic" style="width: 350" value="<%= SiteDetailPic %>">
      <input type="button" value="上传图片" onClick="showUploadDialog('image', 'editForm.SiteDetailPic', '')"></td>
    </tr>
    <tr height="35">
      <td align="right" class="forumRow"></td>
      <td class="forumRowHighlight"><input name="submitSaveEdit" type="submit" id="submitSaveEdit" value="保存主参数设置"></td>
    </tr>
  </form>
</table>
<br />
<%
function SaveSiteInfo()
  '多语言循环保存数据
set rsl = server.createobject("adodb.recordset")
sqll="select * from ChinaQJ_Language order by ChinaQJ_Language_Order"
rsl.open sqll,conn,1,1
while(not rsl.eof)
  if len(trim(request.Form("SiteTitle"&rsl("ChinaQJ_Language_File"))))<4 then
	response.write "<script language='JavaScript'>alert('请详细填写您的"&rsl("ChinaQJ_Language_Name")&"网站标题并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("ComName"&rsl("ChinaQJ_Language_File"))))<4 then
	response.write "<script language='JavaScript'>alert('请详细填写您的"&rsl("ChinaQJ_Language_Name")&"公司名称并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Address"&rsl("ChinaQJ_Language_File"))))<4 then
	response.write "<script language='JavaScript'>alert('请详细填写您的公司"&rsl("ChinaQJ_Language_Name")&"地址并保持至少在两个汉字以上！');" & "history.back()" & "</script>"
    response.end
  end if
rsl.movenext
wend
rsl.close
set rsl=nothing
  if len(trim(request.Form("SiteUrl")))<9 then
	response.write "<script language='JavaScript'>alert('请详细填写您的公司网址！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("ZipCode")))<6 then
	response.write "<script language='JavaScript'>alert('请详细填写邮政编码并保持至少在6个字符以上！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Telephone")))<11 then
	response.write "<script language='JavaScript'>alert('请详细填写联系电话并保持至少在11个字符以上！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Fax")))<11 then
	response.write "<script language='JavaScript'>alert('请详细填写传真号码并保持至少在11个字符以上！');" & "history.back()" & "</script>"
    response.end
  end if
  if len(trim(request.Form("Email")))<6 then
	response.write "<script language='JavaScript'>alert('请详细填写电子邮箱地址并保持至少在6个字符以上！');" & "history.back()" & "</script>"
    response.end
  end if
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from ChinaQJ_Site"
  rs.open sql,conn,1,3
  '多语言循环保存数据
set rsl = server.createobject("adodb.recordset")
sqll="select * from ChinaQJ_Language order by ChinaQJ_Language_Order"
rsl.open sqll,conn,1,1
while(not rsl.eof)
  rs("SiteTitle"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("SiteTitle"&rsl("ChinaQJ_Language_File")))
  rs("Keywords"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("Keywords"&rsl("ChinaQJ_Language_File")))
  rs("Descriptions"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("Descriptions"&rsl("ChinaQJ_Language_File")))
  rs("ComName"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("ComName"&rsl("ChinaQJ_Language_File")))
  rs("Address"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("Address"&rsl("ChinaQJ_Language_File")))
  rs("SiteIndexNotice"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("SiteIndexNotice"&rsl("ChinaQJ_Language_File")))
  rs("Video"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("Video"&rsl("ChinaQJ_Language_File")))
  rs("PageBanner"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("PageBanner"&rsl("ChinaQJ_Language_File")))
  rs("PageBannerType"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("PageBannerType"&rsl("ChinaQJ_Language_File")))
  rs("PageBannerWidth"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("PageBannerWidth"&rsl("ChinaQJ_Language_File")))
  rs("PageBannerHeight"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("PageBannerHeight"&rsl("ChinaQJ_Language_File")))
  rs("SiteDetail"&rsl("ChinaQJ_Language_File"))=trim(Request.Form("SiteDetail"&rsl("ChinaQJ_Language_File")))
rsl.movenext
wend
rsl.close
set rsl=nothing
  rs("SiteUrl")=trim(Request.Form("SiteUrl"))
  rs("ZipCode")=trim(Request.Form("ZipCode"))
  rs("Telephone")=trim(Request.Form("Telephone"))
  rs("Telephone2")=trim(Request.Form("Telephone2"))
  rs("Fax")=trim(Request.Form("Fax"))
  rs("Email")=trim(Request.Form("Email"))
  rs("IcpNumber")=trim(Request.Form("IcpNumber"))
  rs("SiteLogo")=trim(Request.Form("SiteLogo"))
  rs("SiteDetailPic")=trim(Request.Form("SiteDetailPic"))
  rs("ChinaQJ_Stat")=trim(Request.Form("ChinaQJ_Stat"))
  rs("Longitude")=trim(Request.Form("Longitude"))
  rs("Latitude")=trim(Request.Form("Latitude"))
  rs("Proportion")=trim(Request.Form("Proportion"))
  rs("GoogleMapKey")=trim(Request.Form("GoogleMapKey"))
  rs("GoogleMapType")=trim(Request.Form("GoogleMapType"))
  rs("GoogleMapHtml")=trim(Request.Form("GoogleMapHtml"))
  if Request.Form("MesViewFlag")=1 then
    rs("MesViewFlag")=Request.Form("MesViewFlag")
    'Conn.execute "alter table ChinaQJ_Message alter column ViewFlagCh bit default 1"
  else
    rs("MesViewFlag")=0
    'Conn.execute "alter table ChinaQJ_Message alter column ViewFlagCh bit default 0"
  end if
  rs.update
  rs.close
  set rs=nothing
  response.write "<script language='javascript'>alert('系统主参数设置成功！');location.replace('SetSite.asp');</script>"
end function

function ViewSiteInfo()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from ChinaQJ_Site"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
	response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>数据库记录读取错误！</font></div>")
    response.end
  else
  '多语言循环拾取数据
set rsl = server.createobject("adodb.recordset")
sqll="select * from ChinaQJ_Language order by ChinaQJ_Language_Order"
rsl.open sqll,conn,1,1
while(not rsl.eof)
  Lanl=rsl("ChinaQJ_Language_File")
  SiteTitle=rs("SiteTitle"&Lanl)
  Keywords=rs("Keywords"&Lanl)
  Descriptions=rs("Descriptions"&Lanl)
  ComName=rs("ComName"&Lanl)
  Address=rs("Address"&Lanl)
  SiteIndexNotice=rs("SiteIndexNotice"&Lanl)
  Video=rs("Video"&Lanl)
  PageBanner=rs("PageBanner"&Lanl)
  PageBannerType=rs("PageBannerType"&Lanl)
  PageBannerWidth=rs("PageBannerWidth"&Lanl)
  PageBannerHeight=rs("PageBannerHeight"&Lanl)
  SiteDetail=rs("SiteDetail"&Lanl)
  execute("SiteTitle"&Lanl&"=SiteTitle")
  execute("Keywords"&Lanl&"=Keywords")
  execute("Descriptions"&Lanl&"=Descriptions")
  execute("ComName"&Lanl&"=ComName")
  execute("Address"&Lanl&"=Address")
  execute("SiteIndexNotice"&Lanl&"=SiteIndexNotice")
  execute("Video"&Lanl&"=Video")
  execute("PageBanner"&Lanl&"=PageBanner")
  execute("PageBannerType"&Lanl&"=PageBannerType")
  execute("PageBannerWidth"&Lanl&"=PageBannerWidth")
  execute("PageBannerHeight"&Lanl&"=PageBannerHeight")
  execute("SiteDetail"&Lanl&"=SiteDetail")
rsl.movenext
wend
rsl.close
set rsl=nothing
    SiteUrl=rs("SiteUrl")
    ZipCode=rs("ZipCode")
    Telephone=rs("Telephone")
	Telephone2=rs("Telephone2")
    Fax=rs("Fax")
    Email=rs("Email")
    IcpNumber=rs("IcpNumber")
	MesViewFlag=rs("MesViewFlag")
	SiteLogo=rs("SiteLogo")
	SiteDetailPic=rs("SiteDetailPic")
	ChinaQJ_Stat=rs("ChinaQJ_Stat")
    Longitude=rs("Longitude")
    Latitude=rs("Latitude")
    Proportion=rs("Proportion")
    GoogleMapKey=rs("GoogleMapKey")
    GoogleMapType=rs("GoogleMapType")
    GoogleMapHtml=rs("GoogleMapHtml")
    rs.close
    set rs=nothing
  end if
End Function
%>