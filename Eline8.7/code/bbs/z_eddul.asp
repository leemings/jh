<%
sub eddulf()
	dim eddi,addmoney,addbank,addstock,sql,addcool,userWealth,userbank,userstock
	dim dbbank,dbStock,connbank,connstock,rsbank,rsstock,rs,msg
	dbbank="Data/e_Bank.asp"
	Set connbank = Server.CreateObject("ADODB.Connection")
	connbank.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbank)
	dbStock="Data/e_Stock.asp"
	Set connstock = Server.CreateObject("ADODB.Connection")
	connstock.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbStock)
	if membername="" or not founduser then
		userWealth=0
		userbank=0
		userstock=0
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from [User] where username='"&membername&"'"
		rs.open sql,conn,1,3
		userWealth=rs("userWealth")
		set rsbank=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"'"
		rsbank.open sql,connbank,1,3
		if not rsbank.eof then
			userbank=rsbank("savemoney")
		else
			userbank=0
		end if
		set rsstock=server.createobject("adodb.recordset")
		sql="select * from KeHu where ZhangHao='"&membername&"'"
		rsstock.open sql,connstock,1,3
		if not rsstock.eof then
			userstock=int(rsstock("ZiJin"))
		else
			userstock=0
		end if
	end if
	response.write "<br><table align=center cellpadding=3 cellspacing=1 class=tableborder1>"
	response.write "<tr align=center><th width=""100%"">"&forum_info(0)&"互动际遇系统</th></tr>"
	response.write "<tr><td width=100% class=tablebody1>"
	response.write "你在本论坛发帖时所遇到的际遇：<br><ul>"
	Randomize
	eddi=Fix(100*Rnd)
	addmoney=0
	addcool=0
	addbank=0
	addstock=0
	select case eddi
	case 1
		addmoney=50
		response.write "<li>一线天向你颁发美梦成真大奖 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 2
		addmoney=10
		response.write "<li>一线天向你赠送开心活跃奖金 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 3
		addmoney=10
		addcool=5
		response.write "<li>伊然向你赠送开心活跃奖金 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 4
		addmoney=10
		addcool=5
		response.write "<li>玫子向你赠送开心活跃奖金 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 5
		addmoney=-1
		addcool=5
		response.write "<li>我说你说版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 6
		addmoney=-1
		addcool=5
		response.write "<li>玩技交流版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 7
		addmoney=-1
		addcool=5
		response.write "<li>知心话题版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 8
		addmoney=-1
		addcool=5
		response.write "<li>门派议事版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 9
		addmoney=-1
		addcool=5
		response.write "<li>江湖公告版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 10
		addmoney=-1
		addcool=5
		response.write "<li>我说你说版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 11
		addmoney=-1
		addcool=5
		response.write "<li>玩技交流版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 12
		addmoney=-1
		addcool=5
		response.write "<li>知心话题版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 13
		addmoney=-1
		addcool=5
		response.write "<li>门派议事版块向你征收版块建设费 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font>！"
	case 14
		addbank=-userbank/2
		response.write "<li>银行倒闭，存款损失 <font color=red>50%</font>！"
	case 15
		addbank=userbank/2
		response.write "<li>银行抽到大奖，存款增加 <font color=red>50%</font>！"
	case 16
		addbank=-userbank*3/10
		response.write "<li>银行遭挤兑，存款损失 <font color=red>30%</font>！"
	case 17
		addbank=userbank*3/10
		response.write "<li>银行奖励储户，存款增加 <font color=red>30%</font>！"
	case 18
		addbank=-userbank/10
		response.write "<li>银行电脑遭黑客攻击，存款损失 <font color=red>10%</font>！"
	case 19
		addbank=userbank/10
		response.write "<li>银行电脑犯病，存款增加 <font color=red>10%</font>！"
	case 20
		addbank=-50
		response.write "<li>银行利息计算错误，存款损失 <font color=red>50</font> 元！"
	case 21
		addbank=50
		response.write "<li>银行利息计算错误，存款增加 <font color=red>50</font> 元！"
	case 22
		addstock=-userstock/2
		response.write "<li>证券所倒闭，资金损失 <font color=red>50%</font>！"
	case 23
		addstock=userstock/2
		response.write "<li>证券所抽到大奖，资金增加 <font color=red>50%</font>！"
	case 24
		addstock=-userstock*3/10
		response.write "<li>证券所遭挤兑，资金损失 <font color=red>30%</font>！"
	case 25
		addstock=userstock*3/10
		response.write "<li>证券所奖励储户，资金增加 <font color=red>30%</font>！"
	case 26
		addstock=-userstock/10
		response.write "<li>证券所电脑遭黑客攻击，资金损失 <font color=red>10%</font>！"
	case 27
		addstock=userstock/10
		response.write "<li>证券所电脑犯病，资金增加 <font color=red>10%</font>！"
	case 28
		addstock=-50
		response.write "<li>证券所利息计算错误，资金损失 <font color=red>50</font> 元！"
	case 29
		addstock=50
		response.write "<li>证券所利息计算错误，资金增加 <font color=red>50</font> 元！"
	case 30
		addmoney=-2
		addcool=3
		response.write "<li>你路见不平，拔刀相助，花掉你 <font color=red>"&abs(addmoney)&"</font> 元的买刀钱，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 31
		addmoney=-3
		response.write "<li>街上遇到有人打架，混乱中你掉了 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 32
		addmoney=-4
		response.write "<li>你在交通局门口乱停车被罚款 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 33
		addmoney=-userWealth/10
		response.write "<li>做生意不得门路，亏了 <font color=red>10%</font> 的本钱！"
	case 34
		addmoney=-userWealth/10
		response.write "<li>不好意思，小偷又光顾你家，带走了你 <font color=red>10%</font> 的财产！（政府再次提醒人们有钱要存银行）"
	case 35
		addmoney=-20
		response.write "<li>这世界竟然有人整天捡到钱，怎么会没人丢钱呢——摸摸口袋，发现你少了 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 36
		addmoney=-30
		addcool=-1
		response.write "<li>朋友聚赌，你输了 <font color=red>"&abs(addmoney)&"</font> 元，魅力值减 <font color=red>"&abs(addcool)&"</font> ！"
	case 37
		addmoney=-10
		response.write "<li>参加游戏手气不好，一分钱没捞到还花了 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 38
		addmoney=-userWealth/5
		response.write "<li>不幸的你遭到暴徒洗劫！掳去了你 <font color=red>1/4</font> 的现金！（政府再次提醒人们有钱要存银行）"
	case 39
		addmoney=-userWealth*0.10
		response.write "<li>流年不利！货币竟然会贬值，你的财产贬了 <font color=red>10%</font> ！"
	case 40
		addmoney=-3
		response.write "<li>为了发表你的文章，你跑了很长的路，花去了乘车费 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 41
		addmoney=-5
		response.write "<li>到影院看《哈利波特5》花去了你入场费 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 42
		addmoney=-10
		addcool=1
		response.write "<li>路遇乞丐缠身，你忍痛施舍了 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 43
		addmoney=-20
		addcool=2
		response.write "<li>途经美容院，你进去大“修”一番后花去 <font color=red>"&abs(addmoney)&"</font> 元，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 44
		addmoney=-15
		response.write "<li>你的跑车有了小小损坏，花去你 <font color=red>"&abs(addmoney)&"</font> 元的修理费！"
	case 45
		addmoney=-30
		response.write "<li>朋友们要你请客，在大饭店大吃一顿后花去了 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 46
		addmoney=-20
		response.write "<li>你不小心病倒，花去了 <font color=red>"&abs(addmoney)&"</font> 元的治疗费！"
	case 47
		addmoney=-userWealth*0.03
		response.write "<li>政府上门征税，收取你 <font color=red>3%</font> 的个人所得税！"
	case 48
		addmoney=-userWealth/10
		response.write "<li>不小心在阴暗的街角遭到打劫，抢走了你 <font color=red>10%</font> 的现金！（政府提醒人们有钱要存银行）"
	case 49
		addmoney=-10
		response.write "<li>真倒霉！你不小心在俱乐部掉了 <font color=red>"&abs(addmoney)&"</font> 元！"
	case 50
		Randomize
		addmoney=Fix(100*Rnd)+1
		response.write "<li>恭喜，你此次发帖获得了 <font color=red>"&addmoney&"</font> 元的随机特别赠送奖！"
	case 51
		addmoney=10
		response.write "<li>恭喜，你此次发帖获得了 <font color=red>"&addmoney&"</font> 元的特别奖励！"
	case 52
		addmoney=userWealth/10
		response.write "<li>财神爷今日光临本坛，大派财运，你的财产增加了 <font color=red>10%</font> ！"
	case 53
		addmoney=6
		response.write "<li>慈善家今日光临本坛，大派利市，你收了 <font color=red>"&addmoney&"</font> 元的红包！"
	case 54
		addmoney=3
		response.write "<li>政府上门慰问，发给你 <font color=red>"&addmoney&"</font> 元的慰问金！"
	case 55
		addmoney=5
		addcool=-2
		response.write "<li>你在街头拐角捡到 <font color=red>"&addmoney&"</font> 元，收入了你的口袋，魅力值减 <font color=red>"&abs(addcool)&"</font> ！"
	case 56
		addmoney=2
		addcool=2
		response.write "<li>你在街头拐角捡到一条珍珠项链，交还了失主，失主送给你 <font color=red>"&addmoney&"</font> 元的酬金！魅力值加 <font color=red>"&addcool&"</font> ！"
	case 57
		addmoney=10
		response.write "<li>狮子流星带来了一块含金陨石，政府决定派发全民，你分得了 <font color=red>"&addmoney&"</font> 元！"
	case 58
		addmoney=8
		response.write "<li>对于你的积极参与，社区决定发给你 <font color=red>"&addmoney&"</font> 元的奖励金！"
	case 59
		addcool=2
		response.write "<li>心仪你的人为你送上鲜花，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 60
		addmoney=10
		response.write "<li>你被大家推选为今日幸运之星，派发 <font color=red>"&addmoney&"</font> 元的奖金！"
	case 61
		addmoney=3
		response.write "<li>天使降临，你发帖所得将变为 <font color=red>"&addmoney&"</font> 元！"
	case 62
		addmoney=7
		response.write "<li>你幸运地得到发帖累积奖金 <font color=red>"&addmoney&"</font> 元！"
	case 63
		addmoney=6
		response.write "<li>你参加了我们组织的狩猎，获得了 <font color=red>"&addmoney&"</font> 元的猎物兑现金！"
	case 64
		addmoney=10
		response.write "<li>由于你长期对本版的支持，版主决定给你 <font color=red>"&addmoney&"</font> 元的随机表扬金！"
	case 65
		addmoney=6
		response.write "<li>由于你长期对本版的支持，版主决定给你 <font color=red>"&addmoney&"</font> 元的随机表扬金！"
	case 66
		addmoney=4
		response.write "<li>由于你长期对本版的支持，版主决定给你 <font color=red>"&addmoney&"</font> 元的随机表扬金！"
	case 67
		addmoney=3
		response.write "<li>由于你长期对本版的支持，版主决定给你 <font color=red>"&addmoney&"</font> 元的随机表扬金！"
	case 68
		addmoney=2
		response.write "<li>由于你长期对本版的支持，版主决定给你 <font color=red>"&addmoney&"</font> 元的随机表扬金！"
	case 69
		addcool=1
		response.write "<li>今日影院派发嘉宾券，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 70
		addmoney=3
		response.write "<li>今日影院派发嘉宾卷，你将入场券变卖后得到 <font color=red>"&addmoney&"</font> 元！"
	case 71
		addmoney=10
		addcool=2
		response.write "<li>无缘无故有名人请你吃饭，还赠送 <font color=red>"&addmoney&"</font> 元，魅力值加 <font color=red>"&addcool&"</font> ！"
	case 72
		addmoney=7
		response.write "<li>你参加社区游戏得到 <font color=red>"&addmoney&"</font> 元的奖金！"
	case 73
		addmoney=5
		response.write "<li>你参加社区游戏得到 <font color=red>"&addmoney&"</font> 元的奖金！"
	case 74
		addmoney=3
		response.write "<li>你参加社区游戏得到 <font color=red>"&addmoney&"</font> 元的奖金！"
	case 75
		addmoney=10
		addcool=-2
		response.write "<li>你与朋友聚赌赢了 <font color=red>"&addmoney&"</font> 元，魅力值减 <font color=red>"&abs(addcool)&"</font> ！"
	case 76
		addmoney=20
		addcool=-1
		response.write "<li>有人捡到钱把你当成失主送给了你，收到 <font color=red>"&addmoney&"</font> 元，魅力值减 <font color=red>"&abs(addcool)&"</font> ！"
	case 77
		addmoney=userWealth/4
		response.write "<li>做生意遇贵人相助，一天净赚 <font color=red>25%</font> ！"
	case 78
		addmoney=1
		response.write "<li>你在马路边，捡到了<font color=red>一分钱</font>，找不到警察叔叔就把它放到了你的口袋里边！"
	case 79
		addmoney=2
		response.write "<li>你参与社区的广告设计得到 <font color=red>"&addmoney&"</font> 元的佣金！"
	case 80
		addmoney=6
		addcool=2
		response.write "<li>你路见不平，拔刀相助，获得 <font color=red>"&addmoney&"</font> 元的酬金，魅力值加 <font color=red>"&addcool&"</font> ！"
	case else
		response.write "<li>你此次发帖未遇到任何际遇，欢迎你再次发帖！"
	end select
	response.write "</ul>　　　　　　谢谢你的参与！<br>"

	if membername="" or not founduser then
		response.write "<a href=login.asp><font color=red>由于你没有正式登录论坛，所以你在论坛的际遇将只是<b>虑梦一场</b>！</font></a>"
	else
		if addmoney<>0 then
			rs("userWealth")=Fix(rs("userWealth")+addmoney)
			if rs("userWealth")<0 then
				rs("userWealth")=0
			end if
		end if
		if addcool<>0 then
			rs("userCP")=Fix(rs("userCP")+addcool)
			if rs("userCP")<0 then
				rs("userCP")=0
			end if
		end if
		rs.update
		rs.close
		set rs=nothing
		if addbank<>0 and userbank<>0 then
			rsbank("savemoney")=Fix(rsbank("savemoney")+addbank)
			if rsbank("savemoney")<0 then
				rsbank("savemoney")=0
			end if
			rsbank.update
		end if
		rsbank.close
		set rsbank=nothing
		if addstock<>0 and userstock<>0 then
			rsstock("zijin")=rsstock("zijin")+addstock
			if rsstock("zijin")<0 then
				rsstock("zongzijin")=rsstock("zongzijin")-rsstock("zijin")
				rsstock("zijin")=0
			else
				rsstock("Zongzijin")=rsstock("zongzijin")+addstock
			end if
			rsstock.update
		end if
		rsstock.close
		set rsstock=nothing
		if addmoney<>0 or addcool<>0 or addbank<>0 or addstock<>0 then
			msg=""
   if addmoney>0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>获得了</font> <font color=red>"&fix(addmoney)&"</font> 元现金"
   elseif addmoney<0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>损失了</font> <font color=red>"&fix(abs(addmoney))&"</font> 元现金"
   else
    msg=msg&"<font color=Teal>现金没受影响！</font>"
   end if
   if addbank>0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>获得了</font> <font color=red>"&fix(addbank)&"</font> 元存款"
   elseif addbank<0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>损失了</font> <font color=red>"&fix(abs(addbank))&"</font> 元存款"
   else
    msg=msg&"<font color=Teal>存款没受影响！</font>"
   end if
   if addstock>0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>获得了</font> <font color=red>"&fix(addstock)&"</font> 元资金"
   elseif addstock<0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>损失了</font> <font color=red>"&fix(abs(addstock))&"</font> 元资金"
   else
    msg=msg&"<font color=Teal>资金没受影响！</font>"
   end if
   if addcool>0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>增加了</font> <font color=red>"&fix(addcool)&"</font> 点魅力值"
   elseif addcool<0 then
    if msg<>"" then msg=msg&"、"
    msg=msg&"<font color=Teal>减少了</font> <font color=red>"&fix(abs(addcool))&"</font> 点魅力值"
   else
    msg=msg&"<font color=Teal>魅力值没受影响！</font>"
   end if
			if msg<>"" then msg=msg&"！"
			response.write "<hr width=90% size=1 color=#AAAAAA><div align=center>你此次际遇旅行共"&msg&"</div>"
		end if
	end if
	response.write "</td></tr></table><br>"
	connbank.close
	Set connbank = Nothing
	connstock.close
	Set connstock = Nothing
end sub
%>