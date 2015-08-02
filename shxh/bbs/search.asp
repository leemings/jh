<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
searchstr=Request("searchstr")
bid=Request("bid")
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
if activepage<1 then activepage=1
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from 版块 where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>论坛名称："&showcname(rst("id"),rst("名称"),rst("说明"))&"</td><td>版主:"
host=split(rst("版主"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("说明")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table><p><table width=100% bgcolor=cccccc border=1 align=center><tr bgcolor=fffe48 align=center><td>&nbsp;</td><td>&nbsp;</td><td>文章标题</td><td>新窗口</td><td>作者</td><td>字[回]</td><td>最后回复</td><td>最后更新时间</td><td>点击</td></tr>"
rst.Close
rst.Open "select * from bbs where (作者 like '%"&searchstr&"%' or 标题 like '%"&searchstr&"%' or 内容 like '%"&searchstr&"%') and 删除=false and 论坛="&bid&" order by 回复时间  desc",conn,1,3
if  rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=226"
bgc="bgcolor=f7f7f7"
searchlist=" "
pagenum=Ba_pagesize
i=0
do while not rst.EOF 
	if rst("父贴")=0  and instr(searchlist," "&rst("id")&" ")=0 and i>=(activepage-1)*pagenum and i<activepage*pagenum then
	i=i+1
	searchlist=searchlist&rst("id")&" "
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
		msg=msg&"<img src=../images/essence.gif></td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"'  onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst("标题"))<15 and datediff("h",rst("回复时间"),nowtime)<1 then 
		msg=msg&rst("标题")&"</a><img src='../images/new.gif'></td><td>"
	elseif len(rst("标题"))<15 and datediff("h",rst("回复时间"),nowtime)>=1 then 
		msg=msg&rst("标题")&"</a></td><td>"
	elseif len(rst("标题"))>=15 and datediff("h",rst("回复时间"),nowtime)>=1 then 	
		msg=msg&left(rst("标题"),13)&"...</a></td><td>"
	else 
		msg=msg&left(rst("标题"),13)&"...</a><img src='../images/new.gif'></td><td >"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='另开新窗口查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif'></a></td><td>"&showun(rst("作者"))&"</td><td>"
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
	elseif rst("父贴")<>0 and instr(searchlist," "&rst("父贴")&" ")=0  and i>=(activepage-1)*pagenum and i<activepage*pagenum then
	i=i+1	
	searchlist=searchlist&rst("父贴")&" "
	if bgc="bgcolor=f7f7f7" then 
		bgc=""
	else
		bgc="bgcolor=f7f7f7"
	end if
	set rst2=server.CreateObject("adodb.recordset")
	rst2.Open "select * from bbs where id="&rst("父贴"),conn
	msg=msg&"<tr "&bgc&"><td>"
	if datediff("d",rst2("回复时间"),nowtime)>=1 and rst2("回复")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closed.gif'></td><td>"
	elseif	datediff("d",rst2("回复时间"),nowtime)>=1 and rst2("回复")>=Ba_hotbbsnum then
		msg=msg&"<img src='../images/hotclosed.gif'></td><td>"
	elseif	datediff("d",rst2("回复时间"),nowtime)<1 and rst2("回复")<Ba_hotbbsnum then
		msg=msg&"<img src='../images/closedb.gif'></td><td>"
	else
		msg=msg&"<img src='../images/hotclosedb.gif'></td><td>"
	end if	
	if rst2("精华")=True then
		msg=msg&"<img src=images/essence.gif></td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"' onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	else
		msg=msg&"&nbsp;</td><td><a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"'  onmouseover="&chr(34)&"window.status='查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">"
	end if
	if len(rst2("标题"))<15 and datediff("h",rst2("回复时间"),nowtime)<1 then 
		msg=msg&rst2("标题")&"</a><img src='../images/new.gif'></td><td>"
	elseif len(rst2("标题"))<15 and datediff("h",rst2("回复时间"),nowtime)>=1 then 
		msg=msg&rst2("标题")&"</a></td><td>"
	elseif len(rst2("标题"))>=15 and datediff("h",rst2("回复时间"),nowtime)>=1 then 	
		msg=msg&left(rst2("标题"),13)&"...</a></td><td>"
	else 
		msg=msg&left(rst2("标题"),13)&"...</a><img src='../images/new.gif'></td><td >"
	end if
	msg=msg&"<a href='showbbs.asp?bid="&bid&"&fid="&rst2("ID")&"' target='_blank' onmouseover="&chr(34)&"window.status='另开新窗口查看该贴及相关回复贴';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&"><img src='../images/newwin.gif'></a></td><td>"&showun(rst2("作者"))&"</td><td>"
	if rst2("回复作者")="--" then
		msg=msg&len(rst2("内容"))&"</td><td>"
	else
		msg=msg&"<font color=FF0000>["&rst2("回复")&"]</td><td>"
	end if
	if rst2("回复作者")="--" then
		msg=msg&"--</td><td>"
	else
		msg=msg&showun(rst2("回复作者"))&"</td><td>"
	end if		
	msg=msg&rst2("回复时间")&"</td><td align=right>"&rst2("点击")&"</td></tr>"
	rst2.Close
	set rst2=nothing
	elseif rst("父贴")=0  and instr(searchlist," "&rst("id")&" ")=0 and (i<(activepage-1)*pagenum or i>=activepage*pagenum) then
	i=i+1
	searchlist=searchlist&rst("id")&" "
	elseif rst("父贴")<>0 and instr(searchlist," "&rst("父贴")&" ")=0  and (i<(activepage-1)*pagenum or i>=activepage*pagenum) then
	i=i+1
	searchlist=searchlist&rst("父贴")&" "
	end if
	rst.MoveNext
loop
if i mod pagenum =0 then
	pagecount=i/pagenum
else
	pagecount=i\pagenum+1
end if	
msg=msg&"</table><form method=post action='search.asp?bid="&bid&"&searchstr="&searchstr&"' id=form1 name=form1><table width=100% bgcolor=cccccc>"
if activepage>1 then
	msg=msg&"<td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&activepage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
	msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if activepage<pagecount then
	msg=msg&"<td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&activepage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='search.asp?bid="&bid&"&searchstr="&searchstr&"&activepage="&pagecount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
	msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=activepage size=4 value='"&activepage&"'>页/共"&pagecount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form>"&Ba_copyright&"</body></html>"

rst.Close
set rst=nothing
conn.close
set conn=nothing

Response.Write msg
%>