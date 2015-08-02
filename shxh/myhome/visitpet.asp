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
rst.Open "select * from pet where username='"&username&"' and option_T<"&nowtimetype&" and exist=false and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
biological=rst("biological")
sinew=rst("sinew")
maxsinew=rst("maxsinew")
cleanliness=rst("cleanliness")
jiankang=rst("health")
kuaile=rst("happy")
encourage_M=rst("encourage_M")
encourage_C=rst("encourage_C")
beftime=rst("option_T")
decrease=datediff("d",beftime,nowtime)
if decrease=0 then
	msg=biological&"对"&username&"说:又见到你我真开心,我现在很快乐希望你也是"
elseif decrease=1 then
	select case encourage_M
		case "银两","积分"
			conn.Execute "update 用户 set "&encourage_M&"="&encourage_M&"+"&encourage_C&" where 姓名='"&username&"'"
		case "物品"
			rst.Close 
			rst.Open "select * from 物品 where 所有者='"&username&"' and 名称='"&encourage_C&"'",conn
			if rst.EOF or rst.BOF then
				rst.close 
				rst.Open "select * from 商店 where 名称='"&encourage_C&"'",conn
				conn.Execute "insert into 物品(名称,属性,攻击,防御,体力,内力,特效,数量,价格,所有者,寄售,装备) values('"&rst("名称")&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")\2&",'"&username&"',false,false)"
			else
				conn.Execute "update 物品 set 数量=数量+1 where 所有者='"&username&"' and 名称='"&encourage_C&"'"
			end if				
	end select
	conn.Execute  "update pet set sinew="&maxsinew&",health=100,happy=100,cleanliness=100,option_T="&nowtimetype&" where id="&pid
	msg=biological&"对"&username&"说:真希望能天天见到你!这是我送给你的礼物"&encourage_M&encourage_C
else	
	conn.Execute "update pet set sinew=sinew-"&decrease&",health=health-"&decrease&",happy=happy-"&decrease&",cleanliness=cleanliness-"&decrease&",option_T="&nowtimetype&" where id="&pid
	msg=biological&"对"&username&"说:希望明天还能见到你!"
end if
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
<p align=center><font color=0000ff size=5>宠物之家</font><br>这是心的呼唤,这是爱的交流</p><hr>
<p align=center>
<%=msg%><br>
<input type=button value='返回' onclick=javascript:location.href='pet.asp';>
</p>
</body>
</html>