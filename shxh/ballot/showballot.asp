<%
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
pid=Request.QueryString("pid")
if pid="" then Response.Redirect "../error.asp?id=048"
if not isnumeric(pid) then	Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select id,title,explain,deadline,condition from ballotsystem where id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
title=rst("title")
explain=replace(rst("explain"),chr(13)&chr(10),"<br>")
rst.Close
rst.Open "select id,parti,vote from ballot where pid="&pid&" order by vote desc",conn
if rst.EOF or rst.BOF then
	msg="<tr><td>��Ч��ͶƱ��</td></tr><tr><td align=center><input type=button value=' �� �� ' onclick=javascript:location.href='../ballot.asp'  ></td></tr>"
else
	msg="<tr><td align=center bgcolor=ffff00 colspan=2>"&title&"</td></tr><tr><td colspan=2>"&explain&"</td></tr><tr></tr><tr align=center bgcolor=ffff00><td>ѡ��</td><td>��Ʊ</td></tr>"
	do while not rst.EOF
		id=rst("id")
		parti=rst("parti")
		vote=rst("vote")
		msg=msg&"<tr><td>"&parti&"</td><td align=right>"&vote&"</td></tr>"
		rst.MoveNext
	loop
	msg=msg&"<tr align=center><td colspan=2> <input type=button onclick=javascript:location.href='ballot.asp'; value=' �� �� ' > </td></tr></td></tr>"
end if	
rst.Close 
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center><font color=0000ff size=6 face='����'>ͶƱϵͳ</font></p><hr>
<table width='95%' border=1 bordercolorlight=000000 cellspacing=0 cellpadding=2 bordercolordark=FFFFFF align=center>
<%=msg%>
</table>
</html>