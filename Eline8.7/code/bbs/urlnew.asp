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
	set rs=server.createobject("adodb.recordset")
	sql="select * from [fav] where username='"&membername&"'"
	rs.open sql,connurl,1,1
	if rs.eof and rs.bof then
		rs.close
		sql="select * from [fav] where (id is null)"
		rs.open sql,connurl,1,3
		rs.addnew
		rs("username")=membername
		rs("groups")="ѧϰ|����|����|��Դ|��Ϣ"
		rs("groupnum")=5
		rs.update
	end if
	response.redirect "url.asp"
end sub
%>