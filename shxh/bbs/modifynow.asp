<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
bid=Request.Form("bid")
id=Request.Form("id")
title=server.htmlencode(Trim(Request.Form("title")))
content=server.HTMLEncode(TRIM(Request.Form("content")))
if not (isnumeric(bid) and isnumeric(id))then Response.Redirect "../error.asp?id=220"
if username="" then Response.Redirect "../error.asp?id=016"
if title="" then Response.Redirect "../error.asp?id=211"
if content="" then Response.Redirect "../error.asp?id=212"
title=replace(title,"'",chr(34))
content=replace(content,"'",chr(34))
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select ���� from bbs where id="&id&" and ����='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=223"
rst.Close
set rst=nothing
conn.execute "update bbs set ����='"&title&"',����='"&content&"',�ظ�ʱ��="&nowtimetype&" ,�༭=True where id="&id&" and ����='"&username&"'"
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('window.close();',3000);</script></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>��Ϣ�ɹ����ͣ������Ӻ��Զ��ر�</font></td></tr><tr align=center><td><a href='#' onclick='javascript:window.close();'>�ر�</a></td></tr></table></body></html>