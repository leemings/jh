<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
onlinenum=Application("Ba_jxqy_allonlinenum")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if Application("Ba_jxqy_fellow")=true then msg="会员功能开启，目前只有会员能申请保护。"
username=session("Ba_jxqy_username")
chatroomsn=session("Ba_jxqy_userchatroomsn")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
if username="" then Response.Redirect "../error.asp?id=016"
howminute=Request.Form("howminute")
if not isnumeric(howminute) or howminute="" then Response.Redirect "../error.asp?id=024"
howminute=clng(howminute)
if instr(";10;30;60;180;720;1440;",";"&howminute&";")=0 then Response.Redirect "../error.asp?id=024"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
if Application("Ba_jxqy_fellow")=true then
	rst.Open "select protect from 用户 where 姓名='"&username&"' and 会员=true" ,conn
else
	rst.Open "select protect from 用户 where 姓名='"&username&"'" ,conn
end if
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
timetmp=rst("protect")
if timetmp="" or timetmp<nowtime or not isdate(timetmp)  then timetmp=nowtime
if dateadd("n",howminute,timetmp)>dateadd("d",1,nowtime) then Response.Redirect "../error.asp?id=068"
timetmp=dateadd("n",howminute,timetmp)
timetmptype="#"&month(timetmp)&"/"&day(timetmp)&"/"&year(timetmp)&" "&hour(timetmp)&":"&minute(timetmp)&":"&second(timetmp)&"#"
rst.Close
set rst=nothing
conn.Execute "update 用户 set protect="&timetmptype&" where 姓名='"&username&"'"
conn.close
set conn=nothing
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
申请保护成功</font></p>
你在<%=timetmp%>前将不会受人攻击但是也不能攻击别人
<p align=center><input type=button value=' 返 回 ' onclick=javascript:location.href='protect.asp'></p>
</body>
</html>