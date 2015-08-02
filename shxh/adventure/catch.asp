<%
username=session("Ba_jxqy_username")
biological=session("Ba_jxqy_userfight")
if biological="" or biological="none" then
	msg="<FONT color=#ff0000>【错误】</FONT>没有生物可供你捕捉<br>"
elseif username="" then	
	msg="<FONT color=#ff0000>【错误】</FONT>你没有登录或超时断开连接,请重新进入<br>"
else
	randomize()
	nowtime=now()
	nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	if Application("Ba_jxqy_fellow")=true then
		rst.Open "select 体力,防御 from 用户 where 姓名='"&username&"' and 精力>=10 and 会员=true",conn
	else
		rst.Open "select 体力,防御 from 用户 where 姓名='"&username&"' and 精力>=10",conn
	end if
	if rst.EOF or rst.BOF then
		msg="<FONT color=#ff0000>【错误】</FONT>精力不足,请休息一会儿再来试试你的爱心<br>"
	else	
		fhp=rst("体力")
		fdefence=rst("防御")
		rst.Close
		rst.Open "select * from biological where biological='"&biological&"'",conn
		maxsinew=rst("hp")
		encourage=rst("encourage")
		encourage_m=Mid(encourage,1,2)
		encourage_c=Mid(encourage,4)
		attack=rst("attack")-fdefence
		if attack>10000 then
			attack=10000-clng(rnd()*100)
		elseif attack<100 then
			attack=clng(rnd()*1000)
		end if
		fhp=fhp-attack
		if fhp<=0 then
			conn.Execute "update 用户 set 状态='死亡',最后登录时间="&nowtimetype&" where 姓名='"&username&"'"
			session.Abandon
			Response.Write "<script language=javascript>top.location.replace('../error.asp?id=054');</script>"
			msg="<FONT color=#ff0000>【错误】</FONT>你被"&biological&"打死了!<br>"
		else
			conn.Execute "update 用户 set 体力="&fhp&",防御=防御+1,精力=精力-10 where 姓名='"&username&"'"
			rndcatch=rnd()*99+1
			if rndcatch<15 then
				rst.Close
				rst.Open "select * from pet where username='"&username&"' and exist=true"
				if rst.EOF or rst.BOF then
					conn.Execute "insert into pet(username,biological,sinew,maxsinew,cleanliness,health,happy,compete,exist,encourage_m,encourage_c,option_m,option_t) values('"&username&"','"&biological&"',"&maxsinew*rndcatch\100&","&maxsinew&",10,10,10,false,True,'"&encourage_m&"','"&encourage_c&"','无',"&nowtimetype&")"
					msg="<FONT color=#ff0000>【捕捉】</FONT>捕捉成功,请以后好好抚养!"&biological&"的反击使您的生命下降"&attack&"<br>"	
				else
					msg="<FONT color=#ff0000>【捕捉】</FONT>捕捉失败,您无法同时拥有两个宠物!"&biological&"的反击使您的生命下降"&attack&"<br>"	
				end if
			else
				msg="<FONT color=#ff0000>【捕捉】</FONT>捕捉失败,"&biological&"的反击使您的生命下降"&attack&"<br>"	
			end if
		end if	
	end if
	rst.Close
	set rst=nothing
	conn.Close
	set conn=nothing
end if
Response.Write "<script language=javascript>parent.msgfrm.document.writeln('"&msg&"');</script>"
%>