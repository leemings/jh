<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.QueryString("id")
if not isnumeric(pid) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and option_T<"&nowtimetype&" and exist=true and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
msg="<tr><td><input type=hidden name=id value='"&pid&"'>名字</td><td><input type=text name=biological value='"&Trim(rst("biological"))&"' size=14 maxlength=7></td></tr>"
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<title>宠物之家</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<form action=attend2.asp method=post>
<p align=center><font color=0000ff size=5>宠物之家</font><br>这是心的呼唤,这是爱的交流</p><hr>在这儿,你可以给你的宠物取个自己喜欢的名字
<table width=50% align=center border=3>
<%=msg%>
<tr><td>操作</td><td><select name=option>
	<option value='喂食' selected>喂食</option>
	<option value='散步'>散步</option>
	<option value='看病'>看病</option>
	<option value='睡觉'>睡觉</option>
	<option value='洗澡'>洗澡</option>
	</select></td></tr>
<tr><td>时间</td><td><select name='howminute'>
	<option value='10'>10 分 钟</option>
	<option value='30'>半 小 时</option>
	<option value='60'>一 小 时</option>
	<option value='180'>三 小 时</option>
	<option value='720'>十二小时</option>
	<option value='1440'>一&nbsp;&nbsp;&nbsp;&nbsp;天</option>
	</select>
</td></tr>
<tr><td colspan=2 align=center><input type=submit value='照顾'> <input type=button value='返回' onclick=history.back();></td></tr>
</table>
</form>
</body>
</html>