<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
randomize()
rnd0=clng(rnd()*5)+1
rnd1=clng(rnd()*5)+1
rnd2=clng(rnd()*5)+1
if rnd0=rnd1 and rnd0=rnd2 then
ratio=10
point=rnd1*3
grade=3
elseif rnd0=rnd1 then
ratio=2
point=rnd2
grade=2
elseif rnd1=rnd2 then
ratio=2
point=rnd0
grade=2
elseif rnd0=rnd2 then
ratio=2
point=rnd1
grade=2
else
ratio=0
point=0
grade=0
end if
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from dice where username='"&username&"'",conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
zpoint=rst("zpoint")
if ratio<rst("ratio") then ratio=rst("ratio")
zgrade=rst("zgrade")
frequency=rst("frequency")
bet=rst("bet")
rst.Close
set rst=nothing
if point=0 and frequency<3 then
conn.Execute "update dice set frequency=frequency+1 where username='"&username&"'"
msg="庄家"&zpoint&"点;你"&point&"点(第"&frequency&"次).<input type=button value='继续' onclick=javascript:location.href='pdcontinue.asp';>"
elseif 	zgrade>grade or (zgrade=grade and  zpoint>=point) or frequency>2 then
conn.Execute "update dice set z0=6,z1=6,z2=6,zgrade=3,zpoint=18,frequency=0,ratio=3,bet=9999 where username='"&username&"'"
if ratio=10 then
conn.Execute "update 用户 set 银两=银两-"&bet*9&" where 姓名='"&username&"'"
ratio=11
end if
msg="庄家"&zpoint&"点;你"&point&"点(第"&frequency&"次);不好意思你输了"&bet*(ratio-1)&"两银子<input type=button value='重新开始' onclick=javascript:location.href='pdbet.asp';  id=button1 name=button1>"

else
conn.Execute "update dice set z0=6,z1=6,z2=6,zgrade=3,zpoint=18,frequency=0,ratio=3,bet=9999 where username='"&username&"'"
msg="庄家"&zpoint&"点;你"&point&"点(第"&frequency&"次);恭喜你赢了"&bet*(ratio-1)&"两银子<input type=button value='重新开始' onclick=javascript:location.href='pdbet.asp'; >"
conn.Execute "update 用户 set 银两=银两+"&bet*ratio&" where 姓名='"&username&"'"
end if
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../style.css'>
<script language=javascript>
parent.imgfrm.showdice(3,<%=rnd0%>,<%=rnd1%>,<%=rnd2%>);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<p align=center><%=msg%></p>
</html>
