<%
Response.Expires=-1
utime=Request.Form("utime")
if not isnumeric(utime) then Response.Redirect "../error.asp?id=034"
if utime<1 then Response.Redirect "../error.asp?id=035"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"'",conn
if rst.EOF or rst.BOF then session.Abandon
umoney=rst("银两")
rst.Close
set rst=nothing
bill=utime*60
etime=dateadd("h",utime,now())
if umoney<bill then Response.Redirect "../error.asp?id=030"
conn.Execute "update 用户 set 最后登录时间='"&etime&"',状态='眠',体力=体力+"&bill&",银两=银两-"&bill&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
session.Abandon
%>
<head><title>悦来客栈</title><link rel="stylesheet" href="../chatroom/css.css"></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false topmargin=100>
<p align=center><font color=0000FF size=6 face='幼圆'><%=Application("Ba_jxqy_systemname")%>悦来客栈</font></p>
<p><font color=4>承惠纹银<%=bill%>两，非常感谢，我们将会在<%=etime%>之前禁止有人使用你的帐号，欢迎再来！</font></p>
<p align=right><a href="#" onclick=javascript:window.close(); onmouseover="window.status='关闭';return true;" onmouseout="window.status='';return true;">关闭</a></p>
</body>
