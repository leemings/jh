<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
if username="" then  Response.Redirect "../error.asp?id=016"
bid=Request.QueryString("bid")
fid=Request.QueryString("fid")
if not (isnumeric(bid) and isnumeric(fid))then Response.Redirect "../error.asp?id=220"
msg="<html><head><meta http-equiv='Content-Type' content='text/html; charstet=gb2312'><title>"&Ba_pagename&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=0 leftmargin=0 >"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 版块 where id="&bid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=220"
msg=msg&"<table border=0 width=100% bgcolor=cccccc><tr><td>论坛名称："&showcname(rst("id"),rst("名称"),rst("说明"))&"</td><td>版主:"
host=split(rst("版主"),";")
for i=0 to ubound(host)
	msg=msg&showun(host(i))
next
msg=msg&"</td><td width='30%' ><marquee>"&rst("说明")&"</marquee></td><td align=right>"&showhomepage(Ba_homepage,Ba_pagename)&"</td></tr></table>"
rst.Close
rst.Open "select 父贴,标题,内容 from bbs where id="&fid,conn
if rst("父贴")<>0 then fid=rst("父贴")
msg=msg&"<hr><form action=replynow.asp method=post id=form1 name=form1><input type=hidden value='"&bid&"' name=bid><input type=hidden name=fid value="&fid&"><table width=60% border=3 bgcolor=cccccc align=center><tr align=center><td colspan=2><font color=0000ff size=4> 回 复 贴 </font></td></tr><tr><td width='15%'>发贴人</td><td>"&username&"</td></tr><tr><td>●标题</td><td><input type=text name=title size=50 maxlength=50 value='RE:"&rst("标题")&"'></td></tr><tr><td colspan=2>●内容(<a href='ubbcode.htm' target='_blank'>UBB 代码</a>支持)</td></tr><tr><td colspan=2><textarea name=content wrap=PHYSICAL cols=65 rows=10>[quote]"&rst("内容")&"[/quote]"&chr(13)&chr(10)&"</textarea></td></tr><tr align=center><td colspan=2><input type=submit value=' 发 表 ' id=submit1 name=submit1> <input type=reset value=' 重 填 ' id=reset1 name=reset1> <input type=button onclick='javascript:history.back();' value=' 返 回 ' id=button1 name=button1></td></tr></table></form>"&Ba_copyright&"</body></html>"
rst.Close 
set rst=nothing
conn.close
set conn=nothing
Response.Write msg
%>
