<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
bid=Request("bid")
if not isnumeric(bid) or bid="" then Response.Redirect "../error.asp?id=220"
username=session("Ba_jxqy_username")
if instr(Ba_hostname,";"&username&";")=0 then Response.Redirect "../error.asp?id=233"
nowtime=now()
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块 where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>论坛名称："&showcname(rst("id"),rst("名称"),rst("说明"))&"</td><td>版主:"
cname=rst("名称")
host=split(rst("版主"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("说明")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table>"
rst.Close
msg=msg&"<table border=0 bgcolor='"&Ba_bgcolor&"' background='"&Ba_bgimage&"' width='100%'><tr><td>&nbsp;</td><td><font color=00ff00>"&username&"</font>欢迎您的光临</td><td align=right></td></tr></table><table border=1 width=90% align=center bgcolor=cccccc><tr bgcolor=fffe48 align=center><td>&nbsp;</td><td>&nbsp;</td><td>文章标题</td><td>新窗口</td><td>作者</td><td>字[回]</td><td>最后回复</td><td>最后更新时间</td><td>点击</td></tr>"
rst.Open "select * from bbs where 论坛="&bid&" and 删除=true order by 回复时间 desc",conn,1,3
if not(rst.EOF or rst.BOF) then
rst.PageSize=Ba_pagesize
if activepage>rst.PageCount then activepage=rst.PageCount
rst.AbsolutePage=activepage
bgc="bgcolor=f7f7f7"
for i=1 to rst.PageSize
	if rst.EOF or rst.BOF then exit for
	if bgc="bgcolor=f7f7f7" then 
		bgc=""
	else
		bgc="bgcolor=f7f7f7"
	end if	
	msg=msg&"<tr "&bgc&"><td>"
	if datediff("d",rst("回复时间"),nowtime)>=1 and rst("回复")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closed.gif'></td><td>"
	elseif	datediff("d",rst("回复时间"),nowtime)>=1 and rst("回复")>=Ba_hotbbsnum then
		msg=msg&"<img src='../images/hotclosed.gif'></td><td>"
	elseif	datediff("d",rst("回复时间"),nowtime)<1 and rst("回复")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closedb.gif'></td><td>"
	else
		msg=msg&"<img src='../images/hotclosedb.gif'></td><td>"
	end if	
	if rst("精华")=True then
		msg=msg&"<img src='../images/essence.gif'></td><td><a href='showdelbbs.asp?bid="&bid&"&fid="&rst("ID")&"' onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showdelbbs.asp?bid="&bid&"&fid="&rst("ID")&"'  onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst("标题"))<15 and datediff("h",rst("回复时间"),nowtime)<1 then 
		msg=msg&rst("标题")&"</a><img src='../images/new.gif'></td><td align=center>"
	elseif len(rst("标题"))<15 and datediff("h",rst("回复时间"),nowtime)>=1 then 
		msg=msg&rst("标题")&"</a></td><td align=center>"
	elseif len(rst("标题"))>=15 and datediff("h",rst("回复时间"),nowtime)>=1 then 	
		msg=msg&left(rst("标题"),13)&"...</a></td><td align=center>"
	else 
		msg=msg&left(rst("标题"),13)&"...</a><img src='../images/new.gif'></td><td align=center>"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='另开新窗口查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif' border=0></a></td><td>"&showun(rst("作者"))&"</td><td>"
	if rst("回复作者")="--" then
		msg=msg&len(rst("内容"))&"</td><td>"
	else
		msg=msg&"<font color=FF0000>["&rst("回复")&"]</td><td>"
	end if
	if rst("回复作者")="--" then
		msg=msg&"--</td><td>"
	else
		msg=msg&showun(rst("回复作者"))&"</td><td>"
	end if		
	msg=msg&rst("回复时间")&"</td><td align=right>"&rst("点击")&"</td></tr>"
	rst.MoveNext
next
end if
msg=msg&"</table><form method=post action='showdel.asp?bid="&bid&"'><table width=100% bgcolor=cccccc><tr align=center>"
if activepage>1 then
	msg=msg&"<td><a href='showdel.asp?bid="&bid&"&activepage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='showdel.asp?bid="&bid&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
	msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if activepage<rst.PageCount then
	msg=msg&"<td><a href='showdel.asp?bid="&bid&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='showdel.asp?bid="&bid&"&activepage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
	msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=activepage size=4 value='"&activepage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"&Ba_copyright&"</body></html>"
rst.Close
set rst=nothing
conn.close
set conn=nothing
Response.Write msg
%>
