<%
'============================================================
'������ã���̨����˵�����Ŀ
'Website��http://www.sdcms.cn
'Author��ITƽ��
'Date��2008-11-15
'LastUpDate:2011-10
'============================================================
Function Plug_menu
	Plug_menu=""
	Plug_menu=Plug_menu&"<li class=""left_link"" onClick=""DoLocation(this)""><a href=""sdcms_link.asp"" target=""main"">���ӹ���</a>����<a href=""sdcms_vote.asp"" target=""main"">ͶƱ����</a></li>"
	Plug_menu=Plug_menu&"<li class=""left_link"" onClick=""DoLocation(this)""><a href=""sdcms_ad.asp"" target=""main"">������</a>����<a href=""sdcms_comment.asp"" target=""main"">���۹���</a></li>"
	Plug_menu=Plug_menu&"<li class=""left_link"" onClick=""DoLocation(this)""><a href=""sdcms_book.asp"" target=""main"">���Թ���</a>����<a href=""Sdcms_Spider.asp"" target=""main"">֩������</a></li>"
	Plug_menu=Plug_menu&"<li class=""left_link"" onClick=""DoLocation(this)""><a href=""Sdcms_Coll_Item.asp"" target=""main"">�ɼ�����</a></li>"
End Function
%>