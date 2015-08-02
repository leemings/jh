<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allright=Application("Ba_jxqy_allright")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "mid",conn
do while not rst.EOF
	id=rst("id")
	filename=rst("filename")
	songname=rst("songname")
	msg=msg&"<tr><td>"&id&"</td><td>"&filename&"</td><td>"&songname&"</td><td><a href=chgmid.asp?id="&id&"&opt=修改>修改</a> | <a href=chgmid.asp?id="&id&"&opt=删除>删除</a> </td></tr>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>点歌系统</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>文件名</td><td>歌曲名</td><td><a href=chgmid.asp?opt=新增&id=-1>新增歌曲</a></td></tr>
<%=msg%>
</table>
</body>