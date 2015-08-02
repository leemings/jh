<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
pid=Request.QueryString("pid")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
allright=Application("Ba_jxqy_allright")
if not(usergrade=allright and usercorp="官府") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballot where pid="&pid,conn
do while not rst.EOF
	id=rst("id")
	pid=rst("pid")
	parti=rst("parti")
	vote=rst("vote")
	msg=msg&"<tr><td>"&pid&"</td><td>"&parti&"</td><td>"&vote&"</td><td><a href=chgblt.asp?pid="&pid&"&id="&id&"&opt=修改>修改</a> | <a href=chgblt.asp?pid="&pid&"&id="&id&"&opt=删除>删除</a></td></tr>"
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
<p align=center>投票系统</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>投票箱</td><td>选项</td><td>得票</td><%if usergrade=allright and usercorp="官府" then%><td><a href=chgblt.asp?opt=新增&id=-1&pid=<%=pid%>>新增投票选项</a></td><%end if%></tr>
<%=msg%>
</table>
<p align=center><input type=button value='返回' onclick=javascript:location.href='showballot.asp';></p>
</body>