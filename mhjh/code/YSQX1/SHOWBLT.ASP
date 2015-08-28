<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
pid=Request.QueryString("pid")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
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
<body oncontextmenu=self.event.returnValue=false topmargin=0 background='../chatroom/bg1.gif'>
<p align=center>投票系统</p>
<hr>
<table border=3 width=80% align=center>
<tr bgcolor=#FFFF00 align=center><td>投票箱</td><td>选项</td><td>得票</td><%if username=adminname then%><td><a href=chgblt.asp?opt=新增&id=-1&pid=<%=pid%>>新增投票选项</a></td><%end if%></tr>
<%=msg%>
</table>
<p align=center><input type=button value='返回' onclick=javascript:location.href='showballot.asp';></p>
</body>
