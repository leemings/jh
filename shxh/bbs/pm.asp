<!--#include file="set.asp"-->
<%
sendto=Request.QueryString("st")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&sendto&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=228"
rst.Close
set rst=nothing
conn.close
set conn=nothing
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
msg=msg&"<hr><form action=pmnow.asp method=post id=form1 name=form1><table width=80% border=3 bgcolor=cccccc align=center><tr align=center><td colspan=2><font color=0000ff size=4> �� �� </font></td></tr><tr><td>������</td><td>"&username&"</td></tr><tr><td>������</td><td><input type=text name=sendto value='"&sendto&"'></td></tr><tr><td>�����</td><td><input type=text name=title size=40 maxlength=50></td></tr><tr><td colspan=2>������(<a href='ubbcode.htm' target='_blank'>UBB ����</a>֧��)</td></tr><tr><td colspan=2><textarea name=content wrap=PHYSICAL cols=47 rows=10></textarea></td></tr><tr align=center><td colspan=2><input type=submit value=' �� �� ' id=submit1 name=submit1> <input type=reset value=' �� �� ' id=reset1 name=reset1> <input type=button onclick='javascript:window.close();' value=' �� �� ' id=button1 name=button1></td></tr></table></form>"&Ba_copyright&"</body></html>"
Response.Write msg
%>