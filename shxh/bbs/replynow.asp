<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
ip=session("Ba_jxqy_userip")
if username="" then  Response.Redirect "../error.asp?id=016"
bid=Request.Form("bid")
fid=Request.Form("fid")
title=server.htmlencode(Trim(Request.Form("title")))
content=server.HTMLEncode(TRIM(Request.Form("content")))
if not (isnumeric(bid) and isnumeric(fid))then Response.Redirect "../error.asp?id=220"
if title="" then Response.Redirect "../error.asp?id=211"
if content="" then Response.Redirect "../error.asp?id=212"
title=replace(title,"'",chr(34))
content=replace(content,"'",chr(34))
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.execute "insert into bbs(论坛,父贴,标题,内容,点击,回复,作者,写贴时间,回复作者,回复时间,IP,编辑,精华,删除) values("&bid&","&fid&",'"&title&"','"&content&"',0,0,'"&username&"',"&nowtimetype&",'--',"&nowtimetype&",'"&ip&"',False,False,False)"
conn.execute "update bbs set 回复=回复+1,回复作者='"&username&"',回复时间="&nowtimetype&" where id="&fid
conn.execute "update 版块 set 最后作者='"&username&"',最后时间="&nowtimetype&",数量=数量+1 where id="&bid
conn.close
set conn=nothing
%>
<html><head><link rel='stylesheet' href='../chatroom/css.css'><script language=javascript>setTimeout('location.replace("showbbs.asp?bid=<%=bid%>&fid=<%=fid%>");',3000);</script></head><body bgcolor="<%=Ba_bgcolor%>" background="<%=Ba_bgimage%>" topmargin=100><table align=center border=0 width=80% bgcolor=cccccc><tr align=center><td><font color=FF0000 size=4>信息成功发送，三秒钟后自动返回</font></td></tr><tr align=center><td><a href='javascript:location.replace("showbbs.asp?bid=<%=bid%>&fid=<%=fid%>");'>返回</a></td></tr></table></body></html>