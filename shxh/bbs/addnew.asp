<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
ip=session("Ba_jxqy_userip")
if username="" then Response.Redirect "../error.asp?id=016"
bid=Request.Form("bid")
title=server.htmlencode(Trim(Request.Form("title")))
content=server.HTMLEncode(TRIM(Request.Form("content")))
if not isnumeric(bid) then Response.Redirect "../error.asp?id=220"
if title="" then Response.Redirect "../error.asp?id=211"
if content="" then Response.Redirect "../error.asp?id=212"
title=replace(title,"'",chr(34))
content=replace(content,"'",chr(34))
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.createobject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.execute "insert into bbs(��̳,����,����,����,���,�ظ�,����,д��ʱ��,�ظ�����,�ظ�ʱ��,IP,�༭,����,ɾ��) values("&bid&",0,'"&title&"','"&content&"',0,0,'"&username&"',"&nowtimetype&",'--',"&nowtimetype&",'"&ip&"',False,False,False)"
conn.execute "update ��� set �������='"&username&"',���ʱ��="&nowtimetype&",����=����+1 where id="&bid
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('location.replace("showboard.asp?bid=<%=bid%>");',3000);</script></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>��Ϣ�ɹ����ͣ������Ӻ��Զ�����</font></td></tr><tr align=center><td><a href='javascript:location.replace("showboard.asp?bid=<%=bid%>");'>����</a></td></tr></table></body></html>