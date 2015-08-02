<%Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
%>
<!--#include file="set.asp"-->
<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 信件 where 收信人='"&username&"' and 观看=false order by 写信时间 desc",conn
do while not (rst.EOF or rst.BOF)
	content=rst("内容")
	msg=msg&"<table width=100% bgcolor=cccccc border=3><tr><td>发信人："&rst("写信人")&"</td><td>发信时间:"&rst("写信时间")&"</td><td align=right>"&autopro(rst("写信人"))&autopm(rst("写信人"))&"<a href='delpm.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='删除此留言';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"删除此留言"&chr(34)&"><img border=0 src='../images/del.gif'></a></td></tr><tr bgcolor=f7f7f7><td colspan=3>主题："&rst("标题")&"</td><tr align=center><td colspan=3>内容</td></tr><tr bgcolor=f7f7f7><td colspan=3>"&Autolink(content)&"</td></tr></table><br>"
	rst.MoveNext
loop
rst.Close
rst.Open "select * from 信件 where 收信人='"&username&"' and 观看=true order by 写信时间 desc",conn
do while not (rst.EOF or rst.BOF)
	content=rst("内容")
	msg=msg&"<table width=100% bgcolor=dddddd border=1><tr><td>发信人："&rst("写信人")&"</td><td>发信时间:"&rst("写信时间")&"</td><td align=right>"&autopro(rst("写信人"))&autopm(rst("写信人"))&"<a href='delpm.asp?id="&rst("id")&"'onmouseover="&chr(34)&"window.status='删除此留言';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"删除此留言"&chr(34)&"><img border=0 src='../images/del.gif'></a></td></tr><tr bgcolor=d7d7d7><td colspan=3>主题："&rst("标题")&"</td><tr align=center><td colspan=3>内容</td></tr><tr bgcolor=d7d7d7><td colspan=3>"&Autolink(content)&"</td></tr></table><br>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.execute "update 信件 set 观看=True where 收信人='"&username&"'"
conn.close
set conn=nothing
if msg="" then msg="<p align=center><font color=0000FF size=4>您没有留言</font><input type=button value='给朋友送去祝福' onclick=javascript:location.href='../bbs/pm.asp?st="&username&"';></p>"
msg="<html><title>"&Application("Ba_jxqy_systemname")&"</title><link rel='stylesheet' href='../chatroom/css.css'></head><body bgcolor="&Ba_bgcolor&" background="&Ba_bgimage&" oncontextmenu='self.event.returnValue=false' topmargin=20 leftmargin=0>"&msg&"</body></html>"
Response.Write msg
%>