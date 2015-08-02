<!--#include file="../bbs/set.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=233"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "版块",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=224"
do while not rst.EOF
	msg=msg&"<option value='"&rst("id")&"'>"&rst("名称")&"</option>"
	rst.movenext
loop
Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu='self.event.returnValue=false' topmargin=100 leftmargin=0 background='"&Ba_bgimage&"' bgcolor='"&Ba_bgcolor&"'><form action='../bbs/showdel.asp' method=post><table width=60% align=center border=3><tr><td>论坛名称：<select name=bid>"&msg&"</select></td></tr><tr align=center><td><input type=submit value=' 查 看 '></td></tr></table></body></html>"
%>