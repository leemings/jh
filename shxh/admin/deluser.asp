<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
howday=Request.Form("howday")
if not isnumeric(howday) then Response.Redirect "../error.asp?id=024"
howday=clng(howday)
nowtime=now()
abatetime=dateadd("d",-howday,nowtime)
abatetimetype="#"&month(abatetime)&"/"&day(abatetime)&"/"&year(abatetime)&" "&hour(abatetime)&":"&minute(abatetime)&":"&second(abatetime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
conn.Execute "delete  from 用户 where 最后登录时间<"&abatetimetype
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='<%=Application("Ba_jxqy_backgroundimage")%>' bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>'>
业已删除<%=formatdatetime(abatetime,1)%>之后末使用之帐号，但是数据空间不会站即释放，请下载您的数据库(Dafault:system/mxcz.asp)并使用access2000压缩后才能生效！<br>请按返回键返回上一页<p align=center><input type=button onclick=javascript:location.href='selectuser.asp' value='返回'></p>
</body>
</html>