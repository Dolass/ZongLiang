<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
if Instr(session("AdminPurview"),"|52,")=0 then
  response.write ("<br /><br /><div align=""center""><font style=""color:red; font-size:9pt; "")>您没有管理该模块的权限！</font></div>")
  response.end
end if
%>
<link rel="stylesheet" href="Images/Admin_style.css">
<script language="javascript" src="../Scripts/Admin.js"></script>

<br />
<%
set rs = server.createobject("adodb.recordset")
sql="select * from ChinaQJ_Count"
rs.open sql,conn,1,1
if rs.bof and rs.eof then
  response.write("<div style=""color:#F00; text-align:center;"">暂无访问数据！<div>")
else
If Trim(Request.QueryString("Action"))="" Then
%>
<table class="tableBorder" width="80%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th colspan="4">访问数据统计</th>
  </tr>
  <tr height="25">
    <td width="15%" class="forumRow">总访问量：</td>
<%
set rsday = server.createobject("adodb.recordset")
sqlday="select distinct C_hour,C_ip from ChinaQJ_Count"
rsday.open sqlday,conn,1,1
numday=rsday.recordcount
rsday.close
set rsday=nothing

set rsb = server.createobject("adodb.recordset")
sqlb="select top 1 C_time from ChinaQJ_Count order by C_id"
rsb.open sqlb,conn,1,1
strday=rsb(0)
rsb.close
set rsb=nothing

set rsnum = server.createobject("adodb.recordset")
sqlnum="select * from ChinaQJ_Countnum"
rsnum.open sqlnum,conn,1,1
%>
    <td width="35%" class="forumRowHighlight"><%= numday %>/<%= rsnum("N_total") %></td>
    <td width="15%" class="forumRow">开始统计时间：</td>
    <td width="35%" class="forumRowHighlight"><%= strday %></td>
  </tr>
<%
set rsday = server.createobject("adodb.recordset")
sqlday="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&" and C_day="&day(now())&""
rsday.open sqlday,conn,1,1
numday=rsday.recordcount
rsday.close
set rsday=nothing

set rsyday = server.createobject("adodb.recordset")
sqlyday="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&" and C_day="&day(now())-1&""
rsyday.open sqlyday,conn,1,1
numyday=rsyday.recordcount
rsyday.close
set rsyday=nothing
%>
  <tr height="25">
    <td class="forumRow">今日访问量：</td>
    <td class="forumRowHighlight"><%= numday %>/<%= rsnum("N_total") %></td>
    <td class="forumRow">昨日访问量：</td>
    <td class="forumRowHighlight"><%= numyday %>/<%= rsnum("N_yesterday") %></td>
  </tr>
<%
set rsmday = server.createobject("adodb.recordset")
sqlmday="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&""
rsmday.open sqlmday,conn,1,1
nummday=rsmday.recordcount
rsmday.close
set rsmday=nothing

set rsyday = server.createobject("adodb.recordset")
sqlyday="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&""
rsyday.open sqlyday,conn,1,1
numyday=rsyday.recordcount
rsyday.close
set rsyday=nothing
%>
  <tr height="25">
    <td class="forumRow">本月访问量：</td>
    <td class="forumRowHighlight"><%= nummday %>/<%= rsnum("N_total") %></td>
    <td class="forumRow">今年访问量：</td>
    <td class="forumRowHighlight"><%= numyday %>/<%= rsnum("N_total") %></td>
  </tr>
  <tr height="25">
    <td class="forumRow">总统计时间：</td>
    <td class="forumRowHighlight"><%= now()-strday %>天</td>
    <td class="forumRow">平均访问量：</td>
    <td class="forumRowHighlight"><%= rsnum("N_total")/(now()-strday) %></td>
  </tr>
</table>
<br />
<center><input type="text" value="<script type='text/javascript' src='&lt;%=SysRootDir%&gt;Count/MyCount.Asp' charset='utf-8'></script>" style="width: 60%"></center>
<center><font color="#CC0000">将调用代码插入到前台任意位置，有任何疑问可咨询官方售后客服。</font></center>
<%
rsnum.close
set rsnum=nothing
elseif Trim(Request.QueryString("Action"))="all" Then
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="Admin_Count.asp?Action=all"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

set rs1 = server.createobject("adodb.recordset")
sql1="select * from ChinaQJ_Count order by C_time desc"
rs1.open sql1,conn,1,1

if rs1.eof and rs1.bof then
		response.write "目前共有 0 条记录"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then        
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark        		
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
        	else
	        	currentPage=1        
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"条记录"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
<table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th width="5%" align="left"><strong>编号</strong></th>
    <th align="left"><strong>来源网页</strong></th>
    <th width="15%" align="left"><strong>访问时间</strong></th>
    <th width="10%" align="left"><strong>访问IP</strong></th>
    <th width="15%" align="left"><strong>使用浏览器</strong></th>
    <th width="10%" align="left"><strong>使用操作系统</strong></th>
  </tr>
<%
if request("page")<>"" then
j=Cint(request("page"))*MaxPerPage-MaxPerPage+1
else
j=1
end if
do while not rs1.eof
%>
  <tr height="25">
    <td class="leftRow"><%= j %></td>
    <td class="leftRow"><a href="<%= rs1("C_come") %>" title="<%= rs1("C_come") %>" target="_blank"><%= rs1("C_come") %></a></td>
    <td class="leftRow"><%= rs1("C_time") %></td>
    <td class="leftRow"><%= rs1("C_ip") %></td>
    <td class="leftRow"><%= rs1("C_brower") %></td>
    <td class="leftRow"><%= rs1("C_os") %></td>
  </tr>
<%
rs1.movenext
j=j+1
i=i+1
if i>=MaxPerPage then exit do
loop
rs1.close
set rs1=nothing
%>
</table>
<%
end sub
elseif Trim(Request.QueryString("Action"))="chour" Then
%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th colspan="3">24小时统计</th>
  </tr>
  <tr>
    <td colspan="3" class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/1.gif" border="0" width="10" height="<%= yesterday(1) %>" title="访问人数共<%= yesterday(1) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= yesterday(2) %>" title="访问人数共<%= yesterday(2) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= yesterday(3) %>" title="访问人数共<%= yesterday(3) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= yesterday(4) %>" title="访问人数共<%= yesterday(4) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= yesterday(5) %>" title="访问人数共<%= yesterday(5) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= yesterday(6) %>" title="访问人数共<%= yesterday(6) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= yesterday(7) %>" title="访问人数共<%= yesterday(7) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= yesterday(8) %>" title="访问人数共<%= yesterday(8) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= yesterday(9) %>" title="访问人数共<%= yesterday(9) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= yesterday(10) %>" title="访问人数共<%= yesterday(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= yesterday(11) %>" title="访问人数共<%= yesterday(11) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= yesterday(12) %>" title="访问人数共<%= yesterday(12) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= yesterday(13) %>" title="访问人数共<%= yesterday(13) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/14.gif" border="0" width="10" height="<%= yesterday(14) %>" title="访问人数共<%= yesterday(14) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/15.gif" border="0" width="10" height="<%= yesterday(15) %>" title="访问人数共<%= yesterday(15) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/16.gif" border="0" width="10" height="<%= yesterday(16) %>" title="访问人数共<%= yesterday(16) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/17.gif" border="0" width="10" height="<%= yesterday(17) %>" title="访问人数共<%= yesterday(17) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/18.gif" border="0" width="10" height="<%= yesterday(18) %>" title="访问人数共<%= yesterday(18) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/19.gif" border="0" width="10" height="<%= yesterday(19) %>" title="访问人数共<%= yesterday(19) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/20.gif" border="0" width="10" height="<%= yesterday(20) %>" title="访问人数共<%= yesterday(20) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/21.gif" border="0" width="10" height="<%= yesterday(21) %>" title="访问人数共<%= yesterday(21) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/22.gif" border="0" width="10" height="<%= yesterday(22) %>" title="访问人数共<%= yesterday(22) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/23.gif" border="0" width="10" height="<%= yesterday(23) %>" title="访问人数共<%= yesterday(23) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/24.gif" border="0" width="10" height="<%= yesterday(24) %>" title="访问人数共<%= yesterday(24) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">00</td>
          <td class="centerRow">01</td>
          <td class="centerRow">02</td>
          <td class="centerRow">03</td>
          <td class="centerRow">04</td>
          <td class="centerRow">05</td>
          <td class="centerRow">06</td>
          <td class="centerRow">07</td>
          <td class="centerRow">08</td>
          <td class="centerRow">09</td>
          <td class="centerRow">10</td>
          <td class="centerRow">11</td>
          <td class="centerRow">12</td>
          <td class="centerRow">13</td>
          <td class="centerRow">14</td>
          <td class="centerRow">15</td>
          <td class="centerRow">16</td>
          <td class="centerRow">17</td>
          <td class="centerRow">18</td>
          <td class="centerRow">19</td>
          <td class="centerRow">20</td>
          <td class="centerRow">21</td>
          <td class="centerRow">22</td>
          <td class="centerRow">23</td>
        </tr>
      </table></td>
  </tr>
  <tr height="25">
    <td colspan="3" class="centerRow">最近24小时网站访问数据</td>
  </tr>
  <tr height="25">
    <td colspan="3" class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/1.gif" border="0" width="10" height="<%= today(1) %>" title="访问人数共<%= today(1) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= today(2) %>" title="访问人数共<%= today(2) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= today(3) %>" title="访问人数共<%= today(3) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= today(4) %>" title="访问人数共<%= today(4) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= today(5) %>" title="访问人数共<%= today(5) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= today(6) %>" title="访问人数共<%= today(6) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= today(7) %>" title="访问人数共<%= today(7) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= today(8) %>" title="访问人数共<%= today(8) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= today(9) %>" title="访问人数共<%= today(9) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= today(10) %>" title="访问人数共<%= today(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= today(11) %>" title="访问人数共<%= today(11) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= today(12) %>" title="访问人数共<%= today(12) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= today(13) %>" title="访问人数共<%= today(13) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/14.gif" border="0" width="10" height="<%= today(14) %>" title="访问人数共<%= today(14) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/15.gif" border="0" width="10" height="<%= today(15) %>" title="访问人数共<%= today(15) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/16.gif" border="0" width="10" height="<%= today(16) %>" title="访问人数共<%= today(16) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/17.gif" border="0" width="10" height="<%= today(17) %>" title="访问人数共<%= today(17) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/18.gif" border="0" width="10" height="<%= today(18) %>" title="访问人数共<%= today(18) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/19.gif" border="0" width="10" height="<%= today(19) %>" title="访问人数共<%= today(19) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/20.gif" border="0" width="10" height="<%= today(20) %>" title="访问人数共<%= today(20) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/21.gif" border="0" width="10" height="<%= today(21) %>" title="访问人数共<%= today(21) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/22.gif" border="0" width="10" height="<%= today(22) %>" title="访问人数共<%= today(22) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/23.gif" border="0" width="10" height="<%= today(23) %>" title="访问人数共<%= today(23) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/24.gif" border="0" width="10" height="<%= today(24) %>" title="访问人数共<%= today(24) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">00</td>
          <td class="centerRow">01</td>
          <td class="centerRow">02</td>
          <td class="centerRow">03</td>
          <td class="centerRow">04</td>
          <td class="centerRow">05</td>
          <td class="centerRow">06</td>
          <td class="centerRow">07</td>
          <td class="centerRow">08</td>
          <td class="centerRow">09</td>
          <td class="centerRow">10</td>
          <td class="centerRow">11</td>
          <td class="centerRow">12</td>
          <td class="centerRow">13</td>
          <td class="centerRow">14</td>
          <td class="centerRow">15</td>
          <td class="centerRow">16</td>
          <td class="centerRow">17</td>
          <td class="centerRow">18</td>
          <td class="centerRow">19</td>
          <td class="centerRow">20</td>
          <td class="centerRow">21</td>
          <td class="centerRow">22</td>
          <td class="centerRow">23</td>
        </tr>
      </table></td>
  </tr>
  <tr height="25">
    <td colspan="3" class="centerRow">最近24小时网站访问数据</td>
  </tr>
</table>
<%elseif Trim(Request.QueryString("Action"))="cday" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th colspan="3">日统计</th>
  </tr>
  <tr height="25">
    <td colspan="3" class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= Pmonth(1) %>" title="访问人数共<%= Pmonth(1) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= Pmonth(2) %>" title="访问人数共<%= Pmonth(2) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= Pmonth(3) %>" title="访问人数共<%= Pmonth(3) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= Pmonth(4) %>" title="访问人数共<%= Pmonth(4) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= Pmonth(5) %>" title="访问人数共<%= Pmonth(5) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= Pmonth(6) %>" title="访问人数共<%= Pmonth(6) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= Pmonth(7) %>" title="访问人数共<%= Pmonth(7) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= Pmonth(8) %>" title="访问人数共<%= Pmonth(8) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= Pmonth(9) %>" title="访问人数共<%= Pmonth(9) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= Pmonth(10) %>" title="访问人数共<%= Pmonth(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= Pmonth(11) %>" title="访问人数共<%= Pmonth(11) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= Pmonth(12) %>" title="访问人数共<%= Pmonth(12) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/14.gif" border="0" width="10" height="<%= Pmonth(13) %>" title="访问人数共<%= Pmonth(13) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/15.gif" border="0" width="10" height="<%= Pmonth(14) %>" title="访问人数共<%= Pmonth(14) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/16.gif" border="0" width="10" height="<%= Pmonth(15) %>" title="访问人数共<%= Pmonth(15) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/17.gif" border="0" width="10" height="<%= Pmonth(16) %>" title="访问人数共<%= Pmonth(16) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/18.gif" border="0" width="10" height="<%= Pmonth(17) %>" title="访问人数共<%= Pmonth(17) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/19.gif" border="0" width="10" height="<%= Pmonth(18) %>" title="访问人数共<%= Pmonth(18) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/20.gif" border="0" width="10" height="<%= Pmonth(19) %>" title="访问人数共<%= Pmonth(19) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/21.gif" border="0" width="10" height="<%= Pmonth(20) %>" title="访问人数共<%= Pmonth(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/22.gif" border="0" width="10" height="<%= Pmonth(21) %>" title="访问人数共<%= Pmonth(21) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/23.gif" border="0" width="10" height="<%= Pmonth(22) %>" title="访问人数共<%= Pmonth(22) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/24.gif" border="0" width="10" height="<%= Pmonth(23) %>" title="访问人数共<%= Pmonth(23) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/25.gif" border="0" width="10" height="<%= Pmonth(24) %>" title="访问人数共<%= Pmonth(24) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/26.gif" border="0" width="10" height="<%= Pmonth(25) %>" title="访问人数共<%= Pmonth(25) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/27.gif" border="0" width="10" height="<%= Pmonth(26) %>" title="访问人数共<%= Pmonth(26) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/28.gif" border="0" width="10" height="<%= Pmonth(27) %>" title="访问人数共<%= Pmonth(27) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/29.gif" border="0" width="10" height="<%= Pmonth(28) %>" title="访问人数共<%= Pmonth(28) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/30.gif" border="0" width="10" height="<%= Pmonth(29) %>" title="访问人数共<%= Pmonth(29) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/31.gif" border="0" width="10" height="<%= Pmonth(30) %>" title="访问人数共<%= Pmonth(30) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/32.gif" border="0" width="10" height="<%= Pmonth(31) %>" title="访问人数共<%= Pmonth(31) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">01</td>
          <td class="centerRow">02</td>
          <td class="centerRow">03</td>
          <td class="centerRow">04</td>
          <td class="centerRow">05</td>
          <td class="centerRow">06</td>
          <td class="centerRow">07</td>
          <td class="centerRow">08</td>
          <td class="centerRow">09</td>
          <td class="centerRow">10</td>
          <td class="centerRow">11</td>
          <td class="centerRow">12</td>
          <td class="centerRow">13</td>
          <td class="centerRow">14</td>
          <td class="centerRow">15</td>
          <td class="centerRow">16</td>
          <td class="centerRow">17</td>
          <td class="centerRow">18</td>
          <td class="centerRow">19</td>
          <td class="centerRow">20</td>
          <td class="centerRow">21</td>
          <td class="centerRow">22</td>
          <td class="centerRow">23</td>
          <td class="centerRow">24</td>
          <td class="centerRow">25</td>
          <td class="centerRow">26</td>
          <td class="centerRow">27</td>
          <td class="centerRow">28</td>
          <td class="centerRow">29</td>
          <td class="centerRow">30</td>
          <td class="centerRow">31</td>
        </tr>
      </table></td>
  </tr>
  <tr height="25">
    <td class="centerRow" colspan="3">最近一个月的日访问统计数据</td>
  </tr>
  <tr height="25">
    <td class="centerRow" colspan="3"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= Tmonth(1) %>" title="访问人数共<%= Tmonth(1) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= Tmonth(2) %>" title="访问人数共<%= Tmonth(2) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= Tmonth(3) %>" title="访问人数共<%= Tmonth(3) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= Tmonth(4) %>" title="访问人数共<%= Tmonth(4) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= Tmonth(5) %>" title="访问人数共<%= Tmonth(5) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= Tmonth(6) %>" title="访问人数共<%= Tmonth(6) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= Tmonth(7) %>" title="访问人数共<%= Tmonth(7) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= Tmonth(8) %>" title="访问人数共<%= Tmonth(8) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= Tmonth(9) %>" title="访问人数共<%= Tmonth(9) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= Tmonth(10) %>" title="访问人数共<%= Tmonth(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= Tmonth(11) %>" title="访问人数共<%= Tmonth(11) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= Tmonth(12) %>" title="访问人数共<%= Tmonth(12) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/14.gif" border="0" width="10" height="<%= Tmonth(13) %>" title="访问人数共<%= Tmonth(13) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/15.gif" border="0" width="10" height="<%= Tmonth(14) %>" title="访问人数共<%= Tmonth(14) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/16.gif" border="0" width="10" height="<%= Tmonth(15) %>" title="访问人数共<%= Tmonth(15) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/17.gif" border="0" width="10" height="<%= Tmonth(16) %>" title="访问人数共<%= Tmonth(16) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/18.gif" border="0" width="10" height="<%= Tmonth(17) %>" title="访问人数共<%= Tmonth(17) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/19.gif" border="0" width="10" height="<%= Tmonth(18) %>" title="访问人数共<%= Tmonth(18) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/20.gif" border="0" width="10" height="<%= Tmonth(19) %>" title="访问人数共<%= Tmonth(19) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/21.gif" border="0" width="10" height="<%= Tmonth(20) %>" title="访问人数共<%= Tmonth(10) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/22.gif" border="0" width="10" height="<%= Tmonth(21) %>" title="访问人数共<%= Tmonth(21) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/23.gif" border="0" width="10" height="<%= Tmonth(22) %>" title="访问人数共<%= Tmonth(22) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/24.gif" border="0" width="10" height="<%= Tmonth(23) %>" title="访问人数共<%= Tmonth(23) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/25.gif" border="0" width="10" height="<%= Tmonth(24) %>" title="访问人数共<%= Tmonth(24) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/26.gif" border="0" width="10" height="<%= Tmonth(25) %>" title="访问人数共<%= Tmonth(25) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/27.gif" border="0" width="10" height="<%= Tmonth(26) %>" title="访问人数共<%= Tmonth(26) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/28.gif" border="0" width="10" height="<%= Tmonth(27) %>" title="访问人数共<%= Tmonth(27) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/29.gif" border="0" width="10" height="<%= Tmonth(28) %>" title="访问人数共<%= Tmonth(28) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/30.gif" border="0" width="10" height="<%= Tmonth(29) %>" title="访问人数共<%= Tmonth(29) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/31.gif" border="0" width="10" height="<%= Tmonth(30) %>" title="访问人数共<%= Tmonth(30) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/32.gif" border="0" width="10" height="<%= Tmonth(31) %>" title="访问人数共<%= Tmonth(31) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">01</td>
          <td class="centerRow">02</td>
          <td class="centerRow">03</td>
          <td class="centerRow">04</td>
          <td class="centerRow">05</td>
          <td class="centerRow">06</td>
          <td class="centerRow">07</td>
          <td class="centerRow">08</td>
          <td class="centerRow">09</td>
          <td class="centerRow">10</td>
          <td class="centerRow">11</td>
          <td class="centerRow">12</td>
          <td class="centerRow">13</td>
          <td class="centerRow">14</td>
          <td class="centerRow">15</td>
          <td class="centerRow">16</td>
          <td class="centerRow">17</td>
          <td class="centerRow">18</td>
          <td class="centerRow">19</td>
          <td class="centerRow">20</td>
          <td class="centerRow">21</td>
          <td class="centerRow">22</td>
          <td class="centerRow">23</td>
          <td class="centerRow">24</td>
          <td class="centerRow">25</td>
          <td class="centerRow">26</td>
          <td class="centerRow">27</td>
          <td class="centerRow">28</td>
          <td class="centerRow">29</td>
          <td class="centerRow">30</td>
          <td class="centerRow">31</td>
        </tr>
      </table></td>
  </tr>
  <tr height="25">
    <td class="centerRow" colspan="3">最近一个月的日访问统计数据</td>
  </tr>
</table>
<%elseif Trim(Request.QueryString("Action"))="cweek" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th>周统计</th>
  </tr>
  <tr height="25">
    <td class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= Cweek(1) %>" title="访问人数共<%= Cweek(1) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= Cweek(2) %>" title="访问人数共<%= Cweek(2) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= Cweek(3) %>" title="访问人数共<%= Cweek(3) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= Cweek(4) %>" title="访问人数共<%= Cweek(4) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= Cweek(5) %>" title="访问人数共<%= Cweek(5) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= Cweek(6) %>" title="访问人数共<%= Cweek(6) %> 人"></td>
          
          <td width="4%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= Cweek(7) %>" title="访问人数共<%= Cweek(7) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">周日</td>
          <td class="centerRow">周一</td>
          <td class="centerRow">周二</td>
          <td class="centerRow">周三</td>
          <td class="centerRow">周四</td>
          <td class="centerRow">周五</td>
          <td class="centerRow">周六</td>
        </tr>
      </table></td>
  </tr>
</table>
<%elseif Trim(Request.QueryString("Action"))="cmonth" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th>月统计</th>
  </tr>
  <tr height="25">
    <td class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= PCyear(1) %>" title="访问人数共<%= PCyear(1) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= PCyear(2) %>" title="访问人数共<%= PCyear(2) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= PCyear(3) %>" title="访问人数共<%= PCyear(3) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= PCyear(4) %>" title="访问人数共<%= PCyear(4) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= PCyear(5) %>" title="访问人数共<%= PCyear(5) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= PCyear(6) %>" title="访问人数共<%= PCyear(6) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= PCyear(7) %>" title="访问人数共<%= PCyear(7) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= PCyear(8) %>" title="访问人数共<%= PCyear(8) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= PCyear(9) %>" title="访问人数共<%= PCyear(9) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= PCyear(10) %>" title="访问人数共<%= PCyear(10) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= PCyear(11) %>" title="访问人数共<%= PCyear(11) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= PCyear(12) %>" title="访问人数共<%= PCyear(12) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">一月份</td>
          <td class="centerRow">二月份</td>
          <td class="centerRow">三月份</td>
          <td class="centerRow">四月份</td>
          <td class="centerRow">五月份</td>
          <td class="centerRow">六月份</td>
          <td class="centerRow">七月份</td>
          <td class="centerRow">八月份</td>
          <td class="centerRow">九月份</td>
          <td class="centerRow">十月份</td>
          <td class="centerRow">十一月份</td>
          <td class="centerRow">十二月份</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <th>最近一年的所有月份数据统计</th>
  </tr>
  <tr height="25">
    <td class="centerRow"><table class="tableBorder" width="98%" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/2.gif" border="0" width="10" height="<%= Cyear(1) %>" title="访问人数共<%= Cyear(1) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/3.gif" border="0" width="10" height="<%= Cyear(2) %>" title="访问人数共<%= Cyear(2) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/4.gif" border="0" width="10" height="<%= Cyear(3) %>" title="访问人数共<%= Cyear(3) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/5.gif" border="0" width="10" height="<%= Cyear(4) %>" title="访问人数共<%= Cyear(4) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/6.gif" border="0" width="10" height="<%= Cyear(5) %>" title="访问人数共<%= Cyear(5) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/7.gif" border="0" width="10" height="<%= Cyear(6) %>" title="访问人数共<%= Cyear(6) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/8.gif" border="0" width="10" height="<%= Cyear(7) %>" title="访问人数共<%= Cyear(7) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/9.gif" border="0" width="10" height="<%= Cyear(8) %>" title="访问人数共<%= Cyear(8) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/10.gif" border="0" width="10" height="<%= Cyear(9) %>" title="访问人数共<%= Cyear(9) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/11.gif" border="0" width="10" height="<%= Cyear(10) %>" title="访问人数共<%= Cyear(10) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/12.gif" border="0" width="10" height="<%= Cyear(11) %>" title="访问人数共<%= Cyear(11) %> 人"></td>
          
          <td width="5%" class="centerRow" valign="bottom"><img src="Images/13.gif" border="0" width="10" height="<%= Cyear(12) %>" title="访问人数共<%= Cyear(12) %> 人"></td>
          
        </tr>
        <tr>
          <td class="centerRow">一月份</td>
          <td class="centerRow">二月份</td>
          <td class="centerRow">三月份</td>
          <td class="centerRow">四月份</td>
          <td class="centerRow">五月份</td>
          <td class="centerRow">六月份</td>
          <td class="centerRow">七月份</td>
          <td class="centerRow">八月份</td>
          <td class="centerRow">九月份</td>
          <td class="centerRow">十月份</td>
          <td class="centerRow">十一月份</td>
          <td class="centerRow">十二月份</td>
        </tr>
      </table></td>
  </tr>
</table>
<%elseif Trim(Request.QueryString("Action"))="ccome" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th>来源页面</th>
    <th width="30%" colspan="2">所占比例</th>
  </tr>
<%
set rs2 = server.createobject("adodb.recordset")
sql2="select distinct C_come from ChinaQJ_Count"
rs2.open sql2,conn,1,1
set rsnum = server.createobject("adodb.recordset")
sqlnum="select * from ChinaQJ_Countnum"
rsnum.open sqlnum,conn,1,1
do while not rs2.eof
set rs3 = server.createobject("adodb.recordset")
sql3="select * from ChinaQJ_Count where C_come='"&rs2("C_come")&"'"
rs3.open sql3,conn,1,1
numcome=rs3.recordcount
percent=numcome/rsnum("N_total")
percent=round(percent*100,2)
%>
  <tr height="25">
    <td class="leftRow"><a href="<%= rs3("C_come") %>" title="<%= rs3("C_come") %>"><%= rs3("C_come") %></a></td>
    <td width="15%" class="leftRow" valign="middle"><hr style="height: 10; width: <%= percent %>; color: #0000FF"></td>
    <td width="15%" class="leftRow">共 <%= numcome %> 人(<%= percent %>%)</td>
  </tr>
<%
rs3.close
set rs3=nothing
rs2.movenext
loop
rsnum.close
set rsnum=nothing
rs2.close
set rs2=nothing
%>
</table>
<%elseif Trim(Request.QueryString("Action"))="cpage" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th colspan="3">访问页面统计</th>
  </tr>
  <tr height="25">
    <td class="leftRow">访问页面</td>
    <td class="leftRow">访问总数</td>
    <td class="leftRow">所占比例</td>
  </tr>
<%
set rs2 = server.createobject("adodb.recordset")
sql2="select distinct C_page from ChinaQJ_Count"
rs2.open sql2,conn,1,1
set rsnum = server.createobject("adodb.recordset")
sqlnum="select * from ChinaQJ_Countnum"
rsnum.open sqlnum,conn,1,1
do while not rs2.eof
set rs3 = server.createobject("adodb.recordset")
sql3="select * from ChinaQJ_Count where C_page='"&rs2("C_page")&"'"
rs3.open sql3,conn,1,1
if not rs3.bof or not rs3.eof then
numpage=rs3.recordcount
percent=numpage/rsnum("N_total")
percent=round(percent*100,2)
%>
  <tr height="25">
    <td class="leftRow"><a href="<%= rs3("C_page") %>" title="<%= rs3("C_page") %>"><%= rs3("C_page") %></a></td>
    <td width="15%" class="leftRow" valign="middle"><hr style="height: 10; width: <%= percent %>; color: #0000FF"></td>
    <td width="15%" class="leftRow">共 <%= numpage %> 人(<%= percent %>%)</td>
  </tr>
<%
rs3.close
set rs3=nothing
else
Response.Write ""
end if
rs2.movenext
loop
rsnum.close
set rsnum=nothing
rs2.close
set rs2=nothing
%>
</table>
<%elseif Trim(Request.QueryString("Action"))="cip" Then%>
<table class="tableBorder" width="95%" border="0" align="center" cellpadding="5" cellspacing="1">
  <tr height="25">
    <th colspan="3">区域统计</th>
  </tr>
  <tr height="25">
    <td width="20%" class="leftRow">访问地区</td>
    <td class="leftRow">访问总数</td>
    <td width="15%" class="leftRow">所占比例</td>
  </tr>
<%
set rs2 = server.createobject("adodb.recordset")
sql2="select distinct C_where from ChinaQJ_Count"
rs2.open sql2,conn,1,1
set rsnum = server.createobject("adodb.recordset")
sqlnum="select * from ChinaQJ_Countnum"
rsnum.open sqlnum,conn,1,1
do while not rs2.eof
set rs3 = server.createobject("adodb.recordset")
sql3="select * from ChinaQJ_Count where C_where='"&rs2("C_where")&"'"
rs3.open sql3,conn,1,1
numwhere=rs3.recordcount
percent=numwhere/rsnum("N_total")
percent=round(percent*100,2)
hrwidth=int(percent*4)
%>
  <tr height="25">
    <td class="leftRow"><%= rs3("C_where") %></td>
    <td class="leftRow"><hr style="height: 10; width: <%= hrwidth %>; color: #0000FF"></td>
    <td class="leftRow">共 <%= numwhere %> 人(<%= percent %>%)</td>
  </tr>
<%
rs3.close
set rs3=nothing
rs2.movenext
loop
rsnum.close
set rsnum=nothing
rs2.close
set rs2=nothing
%>
</table>
<%
elseif Trim(Request.QueryString("Action"))="del" Then
set rsdel = server.createobject("adodb.recordset")
sqldel="delete from ChinaQJ_Count"
rsdel.open sqldel,conn,3,3
set rsdel=nothing
set rsdel = server.createobject("adodb.recordset")
sqldel="delete from ChinaQJ_Countnum"
rsdel.open sqldel,conn,3,3
set rsdel=nothing
response.redirect "Admin_Count.asp"
End If %>
<%
rs.close
set rs=nothing
End If

Function yesterday(c_hour)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&" and C_day="&day(now())-1&" and C_hour="&C_hour-1&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function today(c_hour)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&" and C_day="&day(now())&" and C_hour="&C_hour-1&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function Pmonth(c_day)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())-1&" and C_day="&C_day&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function Tmonth(c_day)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&month(now())&" and C_day="&C_day&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function Cweek(c_week)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_hour,C_ip from ChinaQJ_Count where C_week="&c_week&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function PCyear(c_month)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())-1&" and C_month="&c_month&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function

Function Cyear(c_month)
dim rs1,sql1
set rs1 = server.createobject("adodb.recordset")
sql1="select distinct C_hour,C_ip from ChinaQJ_Count where C_year="&year(now())&" and C_month="&c_month&""
rs1.open sql1,conn,1,1
numy=rs1.recordcount
rs1.close
set rs1=nothing
Response.Write(numy)
End Function
%>

<br />
