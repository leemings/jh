<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
onlinenum=Application("Ba_jxqy_allonlinenum")
if Application("Ba_jxqy_fellow")=true then msg="会员功能开启，目前只有会员能申请保护。"
username=session("Ba_jxqy_username")
chatroomsn=session("Ba_jxqy_userchatroomsn")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
if username="" then Response.Redirect "../error.asp?id=016"
%>
<html>
<head>
<link rel=stylesheet href='css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=ngimage%>'>
<p align=center>
<font size=5 color=0000ff><%=chatroomname%><br>
<font color=ff0000><%=onlinenum%></font>人在线
<hr>
申请保护</font></p>
<%=msg%>
保护期间，你将无法攻击他人，也不会受他人攻击，也无法解除受保护状态！
<form action= pronow.asp method=post>
<table align=center>
<tr>
	<td>
		<select name='howminute'>
		<option value='10'>10 分 钟</option>
		<option value='30'>半 小 时</option>
		<option value='60'>一 小 时</option>
		<option value='180'>三 小 时</option>
		<option value='720'>十二小时</option>
		<option value='1440'>一&nbsp;&nbsp;&nbsp;&nbsp;天</option>
		</select>
	</td>
</tr>
<tr>
	<td>
		<input type=submit value='保护'>
	</td>
</tr>
</table>
</form>
</html>