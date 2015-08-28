<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
onlinenum=Application("yx8_mhjh_allonlinenum")
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
howminute=Request.Form("howminute")
if not isnumeric(howminute) or howminute="" then Response.Redirect "../error.asp?id=024"
howminute=clng(howminute)
if instr(";10;30;60;180;720;1440;",";"&howminute&";")=0 then Response.Redirect "../error.asp?id=024"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_usercorp")<>"无" then
rst.Open "select protect from 门派 where 门派='"&session("yx8_mhjh_usercorp")&"'" ,conn
else
rst.Open "select protect from 用户 where 门派='"&session("yx8_mhjh_usercorp")&"'" ,conn
end if
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
timetmp=rst("protect")
if timetmp="" or timetmp<nowtime or not isdate(timetmp)  then timetmp=nowtime
if dateadd("n",howminute,timetmp)>dateadd("d",1,nowtime) then Response.Redirect "../error.asp?id=068"
timetmp=dateadd("n",howminute,timetmp)
timetmptype="#"&month(timetmp)&"/"&day(timetmp)&"/"&year(timetmp)&" "&hour(timetmp)&":"&minute(timetmp)&":"&second(timetmp)&"#"
rst.Close
set rst=nothing
conn.Execute "update 门派 set protect="&timetmptype&" where 门派='"&session("yx8_mhjh_usercorp")&"'"
conn.close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='css3.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<p align=center>
<br>
<font color=ff0000><%=onlinenum%></font>人在线
<hr>
申请保护成功</font></p>
你在<%=timetmp%>前，你的门派要塞将不会受人攻击！但时间一过，保护自动消失。
<p align=center><input type=button value=' 返 回 ' onclick=javascript:location.href='protect.asp'></p>
</body>
</html>






