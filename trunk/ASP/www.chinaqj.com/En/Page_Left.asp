﻿<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td class="leftnavtop"><%= ChinaQJLanguageTxt38 %></td>
  </tr>
  <tr>
    <td class="leftnavcenter"><%
if Trim(Request.QueryString("Range"))<>"" then
Range=Trim(Request.QueryString("Range"))
else
Range="Product"
end if
%>
<form id="Search" name="Search" method="get" action="Search.Asp">
          <table width="93%" border="0" align="center" cellpadding="0" cellspacing="0" style="background:url(<%=StylePath%>bg_sear.gif) no-repeat 2px 24px;">
            <tr>
              <td height="28"><input name="Range" type="radio" value="Product" class="inputnoborder" <% If Range="Product" Then %>checked="checked"<% End If %>/>
                <%=ChinaQJLanguageTxt8%>
                <input type="radio" name="Range" value="News" class="inputnoborder" <% If Range="News" Then %>checked="checked"<% End If %>/>
                <%=ChinaQJLanguageTxt9%>
                <input type="radio" name="Range" value="Down" class="inputnoborder" <% If Range="Down" Then %>checked="checked"<% End If %>/>
                <%=ChinaQJLanguageTxt10%></td>
              </tr>
            <tr>
              <td height="28" style="padding-left: 60px"><input name="Keyword" type="text" id="Keyword" style="width:102px;" /></td>
              </tr>
            <tr>
              <td height="28" style="padding-left: 60px"><input type="image" name="imageField3" src="<%=StylePath%>btn_sear_<%=Language%>.gif" class="inputnoborder" /></td>
              </tr>
            <tr>
              <td height="28"><%=ChinaQJLanguageTxt264%>
                <%Call ChinaQJSearchCount(12,18,"")%></td>
              </tr>
            <tr>
              <td><img src="<%=StylePath%>t.gif" width="1" height="9" /></td>
              </tr>
            </table>
          </form></td>
  </tr>
  <tr>
    <td class="leftnavbottom1"></td>
  </tr>
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td class="leftnavtop2"><%= telephone %></td>
  </tr>
  <tr>
    <td class="leftnavcenter"><table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td class="leftcontact leftcontacttop">
            <%= comname %><br />
			<%= ChinaQJLanguageTxt100 %>: <%= address %><br />
			<%= ChinaQJLanguageTxt27 %>: <%= telephone %><br />
			<%= ChinaQJLanguageTxt35 %>: <%= Fax %><br />
			<%= ChinaQJLanguageTxt26 %>: <%= telephone2 %><br />
			<%= ChinaQJLanguageTxt101 %>: <%= email %><br />
			<%= ChinaQJLanguageTxt102 %>: <%= siteurl %>
            </td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td class="leftnavbottom2"></td>
  </tr>
</table>