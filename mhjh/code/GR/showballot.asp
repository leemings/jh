<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
pid=Request.QueryString("pid")
if pid="" then Response.Redirect "../error.asp?id=048"
if not isnumeric(pid) then	Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select id,title,explain,deadline,condition from ballotsystem where id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
title=rst("title")
explain=replace(rst("explain"),chr(13)&chr(10),"<br>")
rst.Close
rst.Open "select id,parti,vote from ballot where pid="&pid&" order by vote desc",conn
if rst.EOF or rst.BOF then
msg="<tr><td>无效的投票箱</td></tr><tr><td align=center><input type=button value=' 返 回 ' onclick=javascript:location.href='ballot.asp'  ></td></tr>"
else
msg="<tr><td align=center bgcolor=005b00 colspan=2>"&title&"</td></tr><tr><td colspan=2>"&explain&"</td></tr><tr></tr><tr align=center bgcolor=005b00><td>选项</td><td>得票</td></tr>"
do while not rst.EOF
id=rst("id")
parti=rst("parti")
vote=rst("vote")
msg=msg&"<tr><td>"&parti&"</td><td align=right>"&vote&"</td></tr>"
rst.MoveNext
loop
msg=msg&"<tr align=center><td colspan=2> <input type=button onclick=javascript:location.href='ballot.asp'; value=' 返 回 ' > </td></tr></td></tr>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel='stylesheet' href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0  background='../chatroom/bg1.gif'>
<p align=center><font face="隶书" size="3" color="#000000">投票系统</font></p><hr>
<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center>
<%=msg%>
</table>
</html>
