<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
bid=Request.QueryString("bid")
id=Request.QueryString("id")
if not (isnumeric(bid) and isnumeric(fid))then Response.Redirect "../error.asp?id=220"
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ��� where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>��̳���ƣ�"&showcname(rst("id"),rst("����"),rst("˵��"))&"</td><td>����:"
host=split(rst("����"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("˵��")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table>"
rst.Close
rst.Open "select ����,����,���� from bbs where id="&id,conn
if username<>rst("����") then Response.Redirect "../error.asp?id=225"
msg=msg&"<hr><form action=modifynow.asp method=post id=form1 name=form1><input type=hidden value='"&bid&"' name=bid><input type=hidden name=id value="&id&"><table width=80% border=3 bgcolor=cccccc align=center><tr align=center><td colspan=2><font color=0000ff size=4> �� �� </font></td></tr><tr><td>������</td><td>"&username&"</td></tr><tr><td>�����</td><td><input type=text name=title size=40 maxlength=50 value='"&rst("����")&"'></td></tr><tr><td colspan=2>������(<a href='ubbcode.htm' target='_blank'>UBB ����</a>֧��)</td></tr><tr><td colspan=2><textarea name=content wrap=PHYSICAL cols=47 rows=10>"&rst("����")&"</textarea></td></tr><tr align=center><td colspan=2><input type=submit value=' �� �� ' id=submit1 name=submit1> <input type=reset value=' �� �� ' id=reset1 name=reset1> <input type=button onclick='javascript:window.close();' value=' �� �� ' id=button1 name=button1></td></tr></table></form>"&Ba_copyright&"</body></html>"
rst.Close 
set rst=nothing
conn.close
set conn=nothing
Response.Write msg
%>