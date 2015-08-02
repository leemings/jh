<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("Ba_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usercorp="官府" and usergrade=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from cardshop order by cprice",conn
do while not rst.EOF
	cname=trim(rst("cname"))
	msg=msg&"<tr><td>"&cname&"</td><td>"&rst("cespecial")&"</td><td>"&rst("ctime")&"</td><td align=right>"&rst("cprice")&"</td><td><a href='upcard1.asp?opt=修改&cid="&rst("id")&"' onmouseover="&chr(34)&"window.status='修改"&cname&"的有关设置';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">修改</a> | <a href='upcard1.asp?opt=删除&cid="&rst("id")&"' onmouseover="&chr(34)&"window.status='删除"&cname&"';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">删除</a></td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<p align=center><font color=0000ff size=4>道具管理</font></p>当会员成能打开时只有会员可以购买和使用道具，反之全部用户可以使用<hr>
<table width=80% border=3 align=center>
<tr bgcolor=ffff00 align=center><td>道具名称</td><td>作用</td><td>作用时限(S)</td><td>价格</td><td><a href='upcard1.asp?opt=新增&cid=-1' onmouseover="window.status='新增卡片道具';return true;" onmouseout="window.status='';return true;">新增</a></td></tr>
<%=msg%>
</table>
</body>
</html>