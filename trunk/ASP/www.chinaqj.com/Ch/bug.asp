<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/NoSQL.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="Function.Asp" -->
<!--#include file="../Include/GetUserInfo.asp" -->
<%
Call SiteInfo
Call FormCheckdata
dim Products
dim mMemID,mRealName,mSex,mCompany,mAddress,mZipCode,mTelephone,mFax,mMobile,mEmail
if session("MemName")<>"" and session("MemLogin")="Succeed" then
  call ChinaQJProductBuyMemInfo()
else
  mSex=ChinaQJLanguageTxt91
  mMemID=0
end If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="shortcut icon" href="favicon.ico"/>
<title>BUG|建议|意见|反馈 - <% =SiteTitle %></title>
<meta name="keywords" content="<% =Keywords %>" />
<meta name="description" content="<% =Descriptions %>" />
<!--#include file="Page_CSS.asp" -->
<link href="Images/global.css" type="text/css" rel="stylesheet">
<script language="javascript" src="../Scripts/Admin.js"></script>
<script type="text/javascript" src="../Scripts/Flash.js"></script>
<script language="JavaScript">
<!--
function killErrors() {
    return true;
}
window.onerror = killErrors;
//-->
</script>
<script src="/Scripts/jquery-1.7.2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	var b = null;
	var url = "<%=GetPrevURL()%>";
	$(document).ready(function(){
		$.get(
			url,
			{},
			function(data){
				b = data;
				var tt = data.match(/<title>(.+)<\/title>/);
				//alert('你要的标题是：'+$('#ttb').text(tt[1]));
				alert(tt[1]);
				//alert("\u7533\u8bf7VID\u5931\u8d25\uff0c\u8bf7\u7a0d\u540e\u518d\u8bd5\uff01");
				//document.getElementById("a_page").innerHTML = GetString(tt[1],10);
				//document.getElementById("a_page").style.display = "block";
				//document.getElementById("div_obj").innerHTML = "您来自:"+tt[1]+"("+url+")页面...";
			}
		);
	});
	function GetString(str,len)
	{
		var strlen = 0; 
		var s = "";
		for(var i = 0;i < str.length;i++)
		{
			if(str.charCodeAt(i) > 128){
				strlen += 2;
			}else{ 
				strlen++;
			}
			s += str.charAt(i);
			if(strlen >= len){ 
				return s ;
			}
		}
		return s;
	}
</script>
</head>

<body>
<% menuname="BUG" %>
<!--#include file="page_Header.asp" -->
<!--#include file="page_BUG.asp" -->
<!--#include file="page_Footer.Asp" -->
</body>
</html>