<%
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.Form("id")
howminute=Request.Form("howminute")
opt=request.form("option")
biological=Request.Form("biological")
if instr(biological,"'")<>0 or instr(biological,chr(34))<>0 or instr(biological,"\")<>0 or instr(biological,"/")<>0 or instr(biological," ")<>0 then Response.Redirect "../error.asp?id=056"
if not(isnumeric(pid) and isnumeric(howminute)) then Response.Redirect "../error.asp?id=024"
if instr(";10;30;60;180;720;1440;",";"&howminute&";")=0 then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
afttime=dateadd("n",howminute,nowtime)
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
afttimetype="#"&month(afttime)&"/"&day(afttime)&"/"&year(afttime)&" "&hour(afttime)&":"&minute(afttime)&":"&second(afttime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and option_T<"&nowtimetype&" and exist=true and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
sinew=rst("sinew")
maxsinew=rst("maxsinew")
cleanliness=rst("cleanliness")
jiankang=rst("health")
kuaile=rst("happy")
beftime=rst("option_T")
decrease=datediff("n",beftime,afttime)\60
sinew=sinew-decrease
cleanliness=cleanliness-decrease
jiankang=jiankang-decrease
kuaile=kuaile-decrease
select case opt
	case "喂食"
		sinew=sinew+howminute
		cleanliness=cleanliness-howminute\100
	case "散步"
		kuaile=kuaile+howminute
		sinew=sinew-howminute\100
	case "看病"
		jiankang=jiankang+howminute
		kuaile=kuaile-howminute\100
	case "睡觉"
		sinew=sinew+howminute
		kuaile=kuaile-howminute\100
	case "洗澡"
		cleanliness=cleanliness+howminute
		sinew=sinew-howminute\100
end select
rst.Close
set rst=nothing
if sinew>maxsinew then 
	sinew=maxsinew
elseif sinew<0 then
	sinew=0
end if		
if jiankang>100 then 
	jiankang=100
elseif jiankang<0 then
	jiankang=0
end if		
if kuaile>100 then 
	kuaile=100
elseif kuaile<0 then
	kuaile=0	
end if
if cleanliness>100 then 
	cleanliness=100
elseif cleanliness<0 then
	cleanliness=0	
end if
conn.Execute "update pet set biological='"&biological&"',sinew="&sinew&",cleanliness="&cleanliness&",health="&jiankang&",happy="&kuaile&",option_T="&afttimetype&",option_M='"&opt&"' where id="&pid
conn.Close
set conn=nothing
%>
<html>
<head>
<title>宠物之家</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000ff size=5>宠物之家</font><br>这是心的呼唤,这是爱的交流</p><hr>当你的宠物的各项属性平均达到80%时可以尝试放生,在放生后,你还可以常去探望他们
<table width=50% align=center border=3>
<tr><td>宠物</td><td><%=biological%></td></tr>
<tr><td>生命</td><td><%Response.Write sinew&"/"&maxsinew%></td></tr>
<tr><td>快乐</td><td><%=kuaile%></td></tr>
<tr><td>健康</td><td><%=jiankang%></td></tr>
<tr><td>清洁</td><td><%=cleanliness%></td></tr>
<tr><td colspan=2 align=center><input type=button value='放生' onclick="javascript:location.href='freepet.asp?id=<%=pid%>';"> <input type=button value='返回' onclick=javascript:location.href='pet.asp';></td></tr>
</table>
</body>
</html>