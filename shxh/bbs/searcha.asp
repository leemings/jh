<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" oncontextmenu='self.event.returnValue=false' topmargin=100 leftmargin=0 ><form action=search.asp method=post><table width=50% bgcolor=cccccc align=center border=3><tr><td>论坛<select name=bid>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=224"
do while not rst.EOF
	Response.Write "<option value='"&rst("id")&"'>"&rst("名称")&"</option>"
rst.MoveNext
loop
rst.close
set rst=nothing
%>
</select></td></tr><tr><td>相关<input type=text size=40 name=searchstr></td></tr><tr align=center><td><input type=submit value=' 搜 寻 '> <input type=reset value=" 重 填 "></td></tr></table></form><%=Ba_copyright%></body></html>