<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
bid=Request.QueryString("bid")
fid=Request.QueryString("fid")
username=session("Ba_jxqy_username")
if instr(Ba_hostname,";"&username&";")=0 then Response.Redirect "../error.asp?id=233"
nowtime=now()
if not (isnumeric(bid) and isnumeric(fid)) then Response.Redirect "../error.asp?id=220"
activepage=Request("activepage")
if activepage="" or not isnumeric(activepage) then activepage=1
activepage=clng(activepage)
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块 where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>论坛名称："&showcname(rst("id"),rst("名称"),rst("说明"))&"</td><td>版主:"
cname=rst("名称")
hostlist=";"&rst("版主")&";"
host=split(rst("版主"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("说明")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table><p><p>"
rst.Close
rst.Open "select tu.门派,tu.等级,tu.签名档,tu.电子邮箱,tb.id,tb.论坛,tb.父贴,tb.标题,tb.内容,tb.点击,tb.回复,tb.作者,tb.写贴时间,tb.回复作者,tb.回复时间,tb.ip,tb.编辑,tb.精华,tb.删除 from bbs tb inner join 用户 tu on tu.姓名=tb.作者 where tb.删除=true and tb.id="&fid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=200"
if isnull(rst("等级")) then 
	msg=msg&"作者资料无法查询，数据出错<br>"
else
	author=rst("作者")
	content=rst("内容")
	sign=rst("签名档")
	msg=msg&"<table border=0 width=95% bgcolor=cccccc align=center><tr><td>文章标题："&rst("标题")&"</td><td align=right>"&autoquote(bid,rst("id"),rst("标题"),content)&autosearch(bid,author)&autopro(author)&autouseremail(rst("电子邮箱"),author)&autopm(author)&automodify(bid,rst("ID"),username,author)&autoreclaim(bid,rst("id"),author,hostlist,username,Ba_hostname)&autoip(rst("ip"),author)&autoessence(bid,rst("id"),hostlist,username,Ba_hostname)&autodelete(bid,rst("id"),author,hostlist,username,Ba_hostname)&"</td></tr><tr><td colspan=2><table width=100% bgcolor=FFFFFF><tr valign=top><td width=100 align=center>"&author&"<br>"&rst("门派")&"<br>"&autolevel(rst("等级"))&"</td><td>"&Autolink(content)&autoedit(rst("编辑"),rst("作者"),rst("回复时间"))&"<br>------------------------<br>"&autolink(sign)&"</td></tr></table></td></tr><tr><td >作者:"&autoemail(rst("电子邮箱"),author)&"</td><td align=right>时间："&rst("写贴时间")&"</td></tr></table><table width=100% bgcolor='"&Ba_bgcolor&"' background='"&Ba_bgimage&"'><tr align=center><td><input type=button onclick=javascript:location.replace('reply.asp?bid="&bid&"&fid="&rst("id")&"'); value=' 回 复 ' ); id=button1 name=button1></td><td><input type=button onclick="&chr(34)&"javascript:location.replace('showboard.asp?bid="&bid&"');"&chr(34)&" value=' 返 回 '  id=button1 name=button1></td></tr></table>"
end if	
rst.Close
rst.Open "select tu.门派,tu.等级,tu.签名档,tu.电子邮箱,tb.id,tb.论坛,tb.父贴,tb.标题,tb.内容,tb.点击,tb.回复,tb.作者,tb.写贴时间,tb.回复作者,tb.回复时间,tb.ip,tb.编辑,tb.精华,tb.删除 from bbs tb inner join 用户 tu on tu.姓名=tb.作者 where tb.删除=true and tb.父贴="&fid&" order by tb.写贴时间",conn
do while not (rst.EOF or rst.BOF)
	if isnull(rst("等级")) then 
		msg=msg&"作者资料无法查询，数据出错<br>"
	else
		author=rst("作者")
		content=rst("内容")
		sign=rst("签名档")
		msg=msg&"<table border=0 width=95% bgcolor=cccccc align=center><tr><td>文章标题："&rst("标题")&"</td><td align=right>"&autoquote(bid,rst("id"),rst("标题"),content)&autosearch(bid,author)&autopro(author)&autouseremail(rst("电子邮箱"),author)&autopm(author)&automodify(bid,rst("ID"),username,author)&autoreclaim(bid,rst("id"),author,hostlist,username,Ba_hostname)&autoip(rst("ip"),author)&autodelete(bid,rst("id"),author,hostlist,username,Ba_hostname)&"</td></tr><tr><td colspan=2><table width=100% bgcolor=FFFFFF><tr valign=top><td width=100 align=center>"&author&"<br>"&rst("门派")&"<br>"&autolevel(rst("等级"))&"</td><td>"&Autolink(content)&autoedit(rst("编辑"),rst("作者"),rst("回复时间"))&"<br>------------------------<br>"&autolink(sign)&"</td></tr></table></td></tr><tr><td >作者:"&autoemail(rst("电子邮箱"),author)&"</td><td align=right>时间："&rst("写贴时间")&"</td></tr></table><br>"
	end if	
	rst.MoveNext 
loop
rst.Close
set rst=nothing
conn.execute "update bbs set 点击=点击+1 where id="&fid
conn.close
set conn=nothing
msg=msg&Ba_copyright&"</body></html>"
Response.Write msg
%>
