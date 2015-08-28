<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
howday=Request.Form("howday")
if not isnumeric(howday) then Response.Redirect "../error.asp?id=024"
howday=clng(howday)
nowtime=now()
abatetime=dateadd("d",-howday,nowtime)
abatetimetype="#"&month(abatetime)&"/"&day(abatetime)&"/"&year(abatetime)&" "&hour(abatetime)&":"&minute(abatetime)&":"&second(abatetime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
conn.Execute "delete  from 用户 where 最后登录时间<"&abatetimetype
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
业已删除<%=formatdatetime(abatetime,1)%>之后末使用之帐号，但是数据空间不会站即释放，请下载您的数据库，并使用access2000压缩后才能生效！<br>请按返回键返回上一页<p align=center><input type=button onclick=javascript:location.href='selectuser.asp' value='返回'></p>
</body>
</html>
