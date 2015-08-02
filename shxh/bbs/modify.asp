<%Response.Expires=-1%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
bid=Request.QueryString("bid")
id=Request.QueryString("id")
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
rst.Open "select 标题,内容,作者 from bbs where id="&id,conn
if username<>rst("作者") then Response.Redirect "../error.asp?id=225"
msg=msg&"<hr><form action=modifynow.asp method=post id=form1 name=form1><input type=hidden value='"&bid&"' name=bid><input type=hidden name=id value="&id&"><table width=80% border=3 bgcolor=cccccc align=center><tr align=center><td colspan=2><font color=0000ff size=4> 修 改 </font></td></tr><tr><td>发贴人</td><td>"&username&"</td></tr><tr><td>●标题</td><td><input type=text name=title size=40 maxlength=50 value='"&rst("标题")&"'></td></tr><tr><td colspan=2>●内容(<a href='ubbcode.htm' target='_blank'>UBB 代码</a>支持)</td></tr><tr><td colspan=2><textarea name=content wrap=PHYSICAL cols=47 rows=10>"&rst("内容")&"</textarea></td></tr><tr align=center><td colspan=2><input type=submit value=' 修 改 ' id=submit1 name=submit1> <input type=reset value=' 重 填 ' id=reset1 name=reset1> <input type=button onclick='javascript:window.close();' value=' 关 闭 ' id=button1 name=button1></td></tr></table></form>"&Ba_copyright&"</body></html>"
rst.Close 
set rst=nothing
conn.close
set conn=nothing
Response.Write msg
%>