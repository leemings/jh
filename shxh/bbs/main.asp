<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块",conn
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu='self.event.returnValue=false' bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&"><table border=3  width=80% bgcolor=cccccc align=center ><tr  align=center><td>论坛名称</td><td>贴子数</td><td>最后发表</td><td>版主</td></tr>"
do while not rst.EOF
	msg=msg&"<tr bgcolor=f7f7f7><td>"&showcname(rst("id"),rst("名称"),rst("说明"))&"<br>"&rst("说明")&"</td><td align=right>"&rst("数量")&"</td><td><a href='#' onclick="&chr(34)&"javascript:window.open('showuser.asp?un="&rst("最后作者")&"','','left=100,top=50,width='+screen.availwidth*2/3+',height='+screen.availheight*3/4+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='查看"&rst("最后作者")&"的个人资料';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"查看"&rst("最后作者")&"的个人资料"&chr(34)&">"&rst("最后作者")&"</a> <br>"&rst("最后时间")&"</td><td>"
	host=split(rst("版主"),";")
	for i=0 to ubound(host)
		msg=msg&showun(host(i))
	next
	msg=msg&"</td></tr>"
	rst.MoveNext
loop 
msg=msg&"</table>"
rst.Close
set rst=nothing
conn.close
set conn=nothing
nowtime=now()
if usergrade=20 and usercorp="官府" then msg=msg&"<p align=center><a href="&chr(34)&"javascript:top.location.replace('../admin/administrator.asp');"&chr(34)&" onmouseover="&chr(34)&"window.status='管理员入口';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"管理员入口"&chr(34)&">管理员入口</a></p>"
msg=msg&Ba_copyright
msg=msg&"</body></html>"
Response.Write msg
%>