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
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "ballotsystem",conn
do while not rst.EOF
	id=rst("id")
	condition=rst("condition")
	active=rst("active")
	if condition="true" then
		condition="所有人"
	elseif condition="false" then
		condition="禁止任何人投票"
	else		
		condition=replace(condition,">=","不小于")
		condition=replace(condition,"<=","不大于")
		condition=replace(condition,">","高于")
		condition=replace(condition,"<","低于")
		condition=replace(condition,"=","为")
		condition=replace(condition," and ","并且")
		condition=replace(condition," or ","或")
	end if
	explain=replace(rst("explain"),chr(13)&chr(10),"<br>")
	if allright=usergrade then
		msg=msg&"<tr><td>"&id&"</td><td>"&rst("title")&"</td><td>"&explain&"</td><td>"&rst("deadline")&"</td><td>"&condition&"</td><td>"&active&"</td><td><a href=chgballot.asp?id="&id&"&opt=修改>修改</a> | <a href=chgballot.asp?id="&id&"&opt=删除>删除</a> | <a href='showblt.asp?pid="&id&"'>查看</a></td></tr>"
	else
		msg=msg&"<tr><td>"&id&"</td><td>"&rst("title")&"</td><td>"&explain&"</td><td>"&rst("deadline")&"</td><td>"&condition&"</td><td>"&active&"</td></tr>"
	end if	
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
<tr bgcolor=#FFFF00 align=center><td>ID</td><td>标题</td><td>说明</td><td>截止日期</td><td>投票条件</td><td>活跃</td><%if usergrade=allright and usercorp="官府" then%><td><a href=chgballot.asp?opt=新增&id=-1>新增投票箱</a></td><%end if%></tr>
<%=msg%>
</table>
</body>