<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="urlconn.asp"-->

<%
dim act
stats="�����ղؼ�"
call nav()
call head_var(2,0,"","")
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ����������ղؼе�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call dvbbs_error()
else
	response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=20><b>�����ղؼ�</b></td><tr><td valign=middle class=tablebody1>"
	call main()
end if

response.write "</table>"
call footer()
%>

<%
sub main()
	dim group,userid
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		response.write "<div align=center>����û���ڱ��������������ղؼУ�����<a href='urlnew.asp'><font color=red>����</font></a>�����ղؼС�</div>"
	else
		group=rs("groups")
		userid=clng(rs("id"))
		response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1 align=center><tr><td width=200 class=tablebody1 valign=top>"
		response.write "<table cellpadding=3 align=center cellspacing=1 class=tableborder1><tr><th>������"
		response.write "<tr><td class=tablebody2><b><img src=pic/ifolder.gif>�����</b>&nbsp;&nbsp;<a href=urlgroup.asp?act=add><font color=red>���</a></font>"
		response.write "<tr><td class=tablebody1>"
		for i=0 to rs("groupnum")-1
			response.write "&nbsp;&nbsp;<b>"&i+1&"��</b>"&trim(split(group,"|")(i))&"&nbsp;&nbsp;&nbsp;<a href=urlgroup.asp?act=edit&name="&trim(split(group,"|")(i))&"><font color=red>����</font></a>&nbsp;<a href=urlgroup.asp?act=del&name="&trim(split(group,"|")(i))&"><font color=red>ɾ��</font></a><br>"
		next
		
		response.write "</select><tr><td cellpadding=3 class=tablebody2><b><a href=urlgroup.asp?act=out><font color=red><img src=pic/delete.gif border=0>ע���ղؼ�</font></a></b>"
		response.write "<tr><td cellpadding=3 class=tablebody1>ע��ɾ���齫ͬʱɾ������������ղ�,ע���ղؼн�ɾ�������ղؼ�,��ͬʱɾ���ղص����е�ַ���޷��ָ���</table>"
		response.write "<form name=form1 method=post action=urlact.asp><td valign=top class=tablebody1 align=left><img src=pic/fav_add1.gif><b>�ҵ������ղؼ�</b>[ȫ��]"
		response.write "<table width=100% cellpadding=3 cellspacing=1 align=center class=tableborder1>"
		dim brs
		for i=0 to rs("groupnum")-1
			response.write "<tr><td class=tablebody2 colspan=3><b>["&split(group,"|")(i)&"]</b>"
			set brs=server.createobject("adodb.recordset")
			sql="select * from [url] where username='"&membername&"' and grouptype='"&split(group,"|")(i)&"'"
			brs.open sql,connurl,1,1
			if brs.eof and brs.bof then
				response.write "<tr><td colspan=3 class=tablebody1>&nbsp;&nbsp;&nbsp;&nbsp;��δ�ղ��κε�ַ��"
			else
				brs.movefirst
				do while not brs.eof
					response.write "<tr><td class=tablebody2 align=center><a href=urlact.asp?act=edit&urlid="&brs("id")&"><font color=blue>"&brs("webname")&"��</font></a><td class=tablebody1><a href="&brs("weburl")&" target=_blank><img src=pic/url.gif border=0>"&brs("weburl")&"</a>"
					if brs("webtext")<>"" then response.write "&nbsp;&nbsp;(˵����"&brs("webtext")&")"
					response.write "<td class=tablebody2 align=center><input type=checkbox name=acter value="&brs("id")&">"
					brs.movenext
				loop
			end if
			brs.close
			set brs=nothing
		next
		response.write "<tr><td class=tablebody2 colspan=2><b><img src=pic/url.gif><font color=red>��ַ����</font></b><td class=tablebody2><input type=checkbox name=chkall value=1>ȫѡ"
		response.write "<tr><td cellpadding=3 cellspacing=1 class=tablebody1 colspan=3 valign=middle>&nbsp;&nbsp;<b><a href=urlact.asp?act=add><font color=blue>���</font></a></b>"
		response.write "&nbsp;&nbsp;<input type=radio name=act value=del><b>ɾ��</b>"
		response.write "&nbsp;&nbsp;<input type=radio name=act value=move><b>�ƶ���</b>����&nbsp;<select name=group>"
		for i=0 to rs("groupnum")-1
			response.write "<option value="&trim(split(group,"|")(i))&">"&trim(split(group,"|")(i))&"</option>"
		next
		response.write "</select>&nbsp;&nbsp;<b>�༭: </b>����ղ���������صı༭��<tr><td class=tablebody2 colspan=3 align=center><input type=submit value=ȷ��>"
		response.write "</table></form></table>"
	end if
end sub
%>