<%
Response.Expires=-1
mycorp=session("Ba_jxqy_usercorp")
mygrade=Session("Ba_jxqy_usergrade")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
maxnosaytime=Application("Ba_jxqy_nosaytime")
username=Session("Ba_jxqy_username")
onlinename=Application("Ba_jxqy_onlinename"&chatroomsn)
lip=Application("Ba_jxqy_lockip")
aberrantname=Application("Ba_jxqy_aberrantname")
msg=server.HTMLEncode(trim(Request.Form("talkmsgtmp")))
namecolor=Request.Form("namecolor")
wordcolor=Request.Form("wordcolor")
sendto=Request.Form("sendto")
expression=Request.Form("expression")
isprivacy=Request.Form("isprivacy")
act=0
if isprivacy="on" and sendto<>"大家" then 
	isprivacy=1
else
	isprivacy=0
end if
if username="" or instr(onlinename,";"&username&";")=0 then Response.Redirect "chaterror.asp?id=000"
if len(msg)>100 then msg=Left(msg,100)
if msg="" or msg="//" or msg="/#" then Response.End
nowtime=now()
if instr(lip,";"&session("Ba_jxqy_userip")&";")<>0 then Response.Write "<script language=javascript>top.location.href="&chr(34)&"chaterror.asp?id=008"&chr(34)&";</script>"
nosaytime=datediff("s",session("Ba_jxqy_userlasttalktime"),nowtime)
if nosaytime>maxnosaytime then 
	Response.Write "<script language=javascript>top.location.href="&chr(34)&"chaterror.asp?id=001"&chr(34)&";</script>"
	Response.End
elseif nosaytime<3 then
	Response.Write "<script language=JavaScript>{alert('有话慢慢说，别噎着了！');}</script>"
	Response.End
end if
if not (sendto="大家" or instr(onlinename,";"&sendto&";")<>0) then
	Response.Write "<script Language=Javascript>alert('"&sendto&"不在聊天室中，不能对其发言。');parent.resfrm.location.href='onlinelist.asp';</script>"
	Response.end
end if
session("Ba_jxqy_userlasttalktime")=nowtime
if instr(aberrantname,username)<>0 then
	abl=Application("Ba_jxqy_aberrantlist")
	ablubd=ubound(abl)
	for i=1 to ablubd step 4
		tt=datediff("s",abl(i+3),nowtime)
		if abl(i)=username and tt<0 and abl(i+2)="麻痹" then
			Response.Write "<script language=javascript>alert('麻痹中，动弹不得！');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="封咒"and left(msg,2)="//" then
			Response.Write "<script language=javascript>alert('封咒状态中，无法使用任何命令！');</script>"
			Response.End
		elseif abl(i)=username and tt<0 and abl(i+2)="疯狂" then
			olt=Application("Ba_jxqy_onlinelist")
			oltubd=ubound(olt)
			randomize()
			rnds=clng((oltubd\8-1)*rnd())
			sendto=olt(rnds*8+1)
		elseif abl(i)=username and tt<0 and instr(";死亡;逮捕;入狱;驱逐;",";"&abl(i+2)&";") then
			Response.Redirect "chaterror.asp?id=000"
		end if	
	next 
end if
if left(msg,2)="//" then
	if chatroomsn=1 or Application("Ba_jxqy_systemname"&chatroomsn)=mycorp or mycorp="官府"  then
	act=Trim(mid(msg,3,3))
	msg=Trim(mid(msg,6))
	if instr(msg,"'")<>0 or (msg="" and  instr(";查看;警告;发钱;逐出;吸星;赠送;偷窃;",";"&act&";")=0) then
		Response.Write "<script language=javascript>alert('"&act&"动作应该带有合适的参数！');</script>"
		Response.End
	elseif instr(onlinename,";"&sendto&";")=0 and instr(";攻击;查看;投掷;送钱;警告;册封;逐出;传功;吸星;赠送;偷窃;道具;",";"&act&";")<>0 then
		Response.Write "<script language=javascript>alert('"&act&"？对谁？？');</script>"
		Response.End
	'elseif sendto=username and instr(";攻击;投掷;送钱;警告;册封;逐出;传功;吸星;赠送;偷窃;道具;",act)<>0 then
	'	Response.Write "<script language=javascript>alert('"&act&"？自己？？');</script>"
	'	Response.End
	end if
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	isprivacy=0
	select case act
		case "查看"
			msg=seeequip(username,sendto)
			isprivacy=1
		case "攻击"
			msg=attack(username,sendto,msg,mygrade)
		case "投掷"
			msg=cast(username,sendto,msg)
		case "送钱"
			msg=sendmoney(username,sendto,msg)
		case "警告"
			msg=alert(username,sendto,mycorp,mygrade,chatroomsn,msg)
		case "发钱"
			msg=distributemoney(username,mycorp,mygrade,chatroomsn)
		case "传言"
			msg="[msg]"&msg&"[/msg]"
		case "公告"
			msg=bulletin(username,sendto,mycorp,mygrade,chatroomsn,msg)	
		case "册封"
			msg=degree(username,sendto,mycorp,msg)
		case "逐出"
			msg=dislodge(username,sendto,mycorp)
		case "传功"
			msg=impart(username,sendto,msg)
		case "吸星"
			msg=burglemp(username,sendto)
		case "偷窃"
			msg=burglemoney(username,sendto)
		case "赠送"
			msg=present(username,sendto,msg)
		case "道具"
			msg=card(username,sendto,msg)	
		case else
			msg="<font color=FF0000>【系统】</font>##，对不起，不明命令，无法执行"	
	end select
	act=2
	set rst=nothing
	conn.Close
	set conn=nothing
	else
		Response.Write "<script language=javascript>alert('你不是本帮弟子，岂能在此撒野！');</script>"
		Response.End
	end if
else
	if left(msg,2)="/#" then
		msg="##<font color="&wordcolor&">"&mid(msg,3)&"</font>"
		act=1
	end if
	randomize()
	rnd1=cint(rnd()*100)
	if rnd1<8 then
		chance=array("体力","内力","攻击","防御","精力","积分","银两","资质","道德")
		chancetxtgood=array("遇少林达摩禅师开坛讲法，##体力增加","武当百岁真人张三丰提拨后进，##内力增加","见血河卫悲回重回人间，突有领悟，##攻击增加","在梦里见到嫦娥舞袖，领悟到神妙步法，##防御增加","##喝醉了酒，却意外见到侠客岛赏善罚恶二使见赐腊八粥，精力增加","途遇老顽童周伯通，##知道了泡分的重要性，积分格外增加","##正在街上行走，观赏MM，却不料有人硬塞给银两","##正在华山后峰挖蚯蚓，做梦也想不到会见到传说中的剑侠风清扬，资质因此而增加","在风波亭北拜谒岳王墓，##道德增加")
		chancetxtbad=array("##长时间熬战聊神谷,以致精力不济，走路也摔跤,体力因此而下降了","##在路上见到了一个美丽的小姑娘，顿起不良之心，不料却被阿紫偷偷地化去了内力","##闭关苦练躺尸剑法，却不得要领，攻击只此而减少","可笑##没事想学学荆无命的自残剑法，最后只落得满身伤疤，防御因此而减少","##闲来无事逛了逛翠花楼，结果可想而知，精力下降了","城隍庙的土地实然睁开了眼，指着##说：你很少来吧，扣你积分以示公正！,虽然##呼天呛地大喊冤枉，积分还是被扣去了","##交友不慎,银两被骗去了","##在练武场认识了薛蟠，大谈武道，资质因此而减少了","##途见空空儿，苦心拜师学艺，不料功夫没有学到，道德却下降了")
		rnd2=cint(rnd()*200)-80
		if rnd2>0 then 
			rndtxt=chancetxtgood(rnd1)&rnd2
		else
			rnd2=rnd2-1
			rndtxt=chancetxtbad(rnd1)&abs(rnd2)
		end if	
		set conn=Server.CreateObject("adodb.connection")
		conn.open Application("Ba_jxqy_connstr")
		conn.execute "update 用户 set "&chance(rnd1)&"="&chance(rnd1)&"+"&rnd2&" where 姓名='"&username&"'"
		conn.close
		set conn=nothing
		talkarr=Application("Ba_jxqy_talkarr")
		talkpoint=clng(Application("Ba_jxqy_talkpoint"))
		dim nta(600)
		j=1
		for i=11 to 600 step 10
			nta(j)=talkarr(i)
			nta(j+1)=talkarr(i+1)
			nta(j+2)=talkarr(i+2)
			nta(j+3)=talkarr(i+3)
			nta(j+4)=talkarr(i+4)
			nta(j+5)=talkarr(i+5)
			nta(j+6)=talkarr(i+6)
			nta(j+7)=talkarr(i+7)
			nta(j+8)=talkarr(i+8)
			nta(j+9)=talkarr(i+9)
			j=j+10
		next
		nta(591)=talkpoint+1
		nta(592)=1
		nta(593)=0
		nta(594)=username
		nta(595)="大家"
		nta(596)=""
		nta(597)=namecolor
		nta(598)=wordcolor
		nta(599)="<font color=FF0000>【运气】</font>"&rndtxt&"<font class=timsty>("&nowtime&")</font>"
		nta(600)=chatroomsn
		Application.Lock
		Application("Ba_jxqy_talkpoint")=talkpoint+1
		Application("Ba_jxqy_talkarr")=nta
		Application.UnLock
		erase nta
		erase chancetxtbad
		erase chancetxtgood
		erase chance
	end if	
end if
msg=replace(msg,"\","\\")
msg=replace(msg,"/","\/")
msg=replace(msg,chr(34),"\"&chr(34))
msg=replace(msg,chr(39),"\"&chr(39))
talkarr=Application("Ba_jxqy_talkarr")
talkpoint=clng(Application("Ba_jxqy_talkpoint"))
dim newtalkarr(600)
j=1
for i=11 to 600 step 10
	newtalkarr(j)=talkarr(i)
	newtalkarr(j+1)=talkarr(i+1)
	newtalkarr(j+2)=talkarr(i+2)
	newtalkarr(j+3)=talkarr(i+3)
	newtalkarr(j+4)=talkarr(i+4)
	newtalkarr(j+5)=talkarr(i+5)
	newtalkarr(j+6)=talkarr(i+6)
	newtalkarr(j+7)=talkarr(i+7)
	newtalkarr(j+8)=talkarr(i+8)
	newtalkarr(j+9)=talkarr(i+9)
	j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=act
newtalkarr(593)=isprivacy
newtalkarr(594)=username
newtalkarr(595)=sendto
newtalkarr(596)=expression
newtalkarr(597)=namecolor
newtalkarr(598)=wordcolor
newtalkarr(599)=msg&"<font class=timsty>（"&time()&"）<\/font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("Ba_jxqy_talkpoint")=talkpoint+1
Application("Ba_jxqy_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
dim showarr()
mytalkpoint=session("Ba_jxqy_usertalkpoint")
newtalkpoint=0
talkarr=Application("Ba_jxqy_talkarr")
j=1
for i=1 to 600 step 10
	newtalkpoint=talkarr(i)
	if newtalkpoint>mytalkpoint and cstr(talkarr(i+9))=cstr(chatroomsn) and (talkarr(i+2)="0" or talkarr(i+4)="大家" or (talkarr(i+2)="1" and (talkarr(i+3)=username or talkarr(i+4)=username))) then
	redim preserve showarr(j),showarr(j+1),showarr(j+2),showarr(j+3),showarr(j+4),showarr(j+5),showarr(j+6),showarr(j+7),showarr(j+8),showarr(j+9)
	showarr(j)=talkarr(i)
	showarr(j+1)=talkarr(i+1)
	showarr(j+2)=talkarr(i+2)
	showarr(j+3)=talkarr(i+3)
	showarr(j+4)=talkarr(i+4)
	showarr(j+5)=talkarr(i+5)
	showarr(j+6)=talkarr(i+6)
	showarr(j+7)=talkarr(i+7)
	showarr(j+8)=talkarr(i+8)
	showarr(j+9)=talkarr(i+9)
	j=j+10
	end if
next
if j>1 then
Response.Write "<script language=javascript>"
	for i=1 to ubound(showarr) step 10
		Response.Write "parent.showmsg('"&showarr(i+1)&"','"&showarr(i+2)&"','"&showarr(i+3)&"','"&showarr(i+4)&"','"&showarr(i+5)&"','"&showarr(i+6)&"','"&showarr(i+7)&"','"&showarr(i+8)&"');"
	next
Response.Write "</script>"
end if
erase talkarr
erase showarr
if newtalkpoint>mytalkpoint then session("Ba_jxqy_usertalkpoint")=newtalkpoint
function attack(un,st,mg,gr)
	timetmp=now()
	rst.Open "select ta.消耗内力,ta.基本攻击+tu.攻击 as 攻击,ta.特效,ta.攻击说明,tu.内力,tu.积分,tu.道德,tu.protect from 攻击 ta inner join 用户 tu ON tu.姓名=ta.姓名 where ta.姓名='"&un&"' and ta.招式='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		attack="<font color=FF0000>【攻击】</font>##咬牙切齿地对%%说：'等我学会了"&mg&",一定会再来找你一比高下！"
	else
		basemp=rst("消耗内力")
		att=rst("攻击")
		especial=rst("特效")
		fmorcal=rst("道德")
		atdeclaration=rst("攻击说明")
		if isnull(especial) then especila="无"
		if rst("内力")<basemp then
			attack="<font color=FF0000>【攻击】</font>##本意是想用"&mg&"攻击%%，可惜心有余而力不足。"
		elseif rst("积分")<Application("Ba_jxqy_newuser") then
			attack="<font color=FF0000>【攻击】</font>##你手无缚鸡之力，还是先别打人了！"
		elseif rst("protect")>timetmp then
			attack="<font color=FF0000>【攻击】</font>##你现在处于保护阶段，无法攻击%%！"
		else
			rst.Close
			rst.Open "select 体力,内力,防御,道德,特技,积分,状态,protect from 用户 where 姓名='"&st&"'",conn
			if rst("积分")<Application("Ba_jxqy_newuser") then
				attack="<font color=FF0000>【攻击】</font>%%还是个孩子，##你怎么能忍心下手呢？"
			elseif rst("protect")>timetmp then
				attack="<font color=FF0000>【攻击】</font>%%现在处于受保护阶段，##你无法使用"&mg&"攻击！"	
			else	
				resatt=att-rst("防御")
				stespecial=rst("特技")
				condition=rst("状态")
				tmorcal=rst("道德")
				Maxattack=Application("Ba_jxqy_Maxattack")*gr\20
				randomize()
				if resatt<0 then 
					resatt=clng(rnd()*100)+1
				elseif resatt>Maxattack then 
					resatt=Maxattack-clng(rnd()*100)
				end if
				morcal=clng(fmorcal-tmorcal*0.01)
				if morcal>=2100000000 then 
					morcal=2100000000
				elseif morcal<=-2100000000 then
					morcal=-2100000000
				end if
				if resatt>rst("体力") and condition<>"死亡" then
					especial="死亡"
					especialtime=1800
					conn.execute "update 用户 set 道德="&morcal&",内力=内力-"&basemp&",攻击=攻击+1 where 姓名='"&un&"'"
					conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
					attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，竟把%%打死了！！"
				elseif 	resatt>rst("体力") and condition="死亡" then
					especial="死亡"
					especialtime=1800
					conn.execute "update 用户 set 道德=道德-1,内力=内力-"&basemp&",攻击=攻击+1 where 姓名='"&un&"'"
					attack="<font color=FF0000>【##攻击】</font>##打死了%%，竞然开始鞭尸，道德减1！！"
				else
					especialtime=resatt\100
					if especialtime<10 then 
						especial="无"
					elseif especialtime>180 then 
						especialtime=180
					end if	
					conn.Execute "update 用户 set 内力=内力-"&basemp&",攻击=攻击+1 where 姓名='"&un&"'"
					conn.Execute "update 用户 set 体力=体力-"&resatt&",防御=防御+1 where 姓名='"&st&"'"
					attack="<font color=FF0000>【##攻击】</font>"&atdeclaration&"，使%%生命减少"&resatt&"。<bgsound src='../mid/dw.wav' loop=1>"
				end if
				if instr(stespecial,";抗"&especial&";")<>0 then especial="无"
				if especial<>"无" then	attack=attack&newaberrant(st,un,especial,especialtime)
			end if
		end if
	end if
	rst.Close
end function
function seeequip(un,st)
	rst.Open "select 门派,身份,等级,配偶 from 用户 where 姓名='"&st&"'",conn
	seeequip="<font color=FF0000>【查看】</font>##得到关于%%的情报。门派："&rst("门派")&"；身份："&rst("身份")&"；配偶："&rst("配偶")&"；"
	if rst("等级")<3 then
		seeequip=seeequip&"看起来象个新手。"
	elseif rst("等级")<6 and  rst("等级")>=3 then
		seeequip=seeequip&"看起来有点儿实力。"
	elseif rst("等级")<9 and  rst("等级")>=6 then
		seeequip=seeequip&"看起来来头不小。"
	else
		seeequip=seeequip&"这种人最好不要惹。"
	end if
	rst.Close
end function
function cast(un,st,mg)
	tt=now()
	rst.Open "select tc.攻击,tc.特效,tu.积分,tu.道德,tu.protect from 物品 tc inner join 用户 tu on tc.所有者=tu.姓名 where tc.名称='"&mg&"' and tc.所有者='"&un&"' and 数量>0",conn
	if rst.EOF or rst.BOF then 
		cast="<font color=FF0000>【投掷】</font>##，你并没有"&mg&"，所以无法投掷！"
	elseif rst("积分")<Application("Ba_jxqy_newuser") then 
		cast="<font color=FF0000>【投掷】</font>##你还是个新人，别惹祸上身！"
	elseif rst("protect")>tt then
		cast="<font color=FF0000>【投掷】</font>##你现在处于保护阶段,无法攻击%%！"
	else
		fmorcal=rst("道德")
		athp=rst("攻击")
		atespecial=rst("特效")
		rst.Close
		rst.Open "select 积分,体力,道德,特技,状态,protect from 用户 where 姓名='"&st&"'",conn
		if rst("积分")<Application("Ba_jxqy_newuser") then 
			cast="<font color=FF0000>【投掷】</font>%%你还是个新人，你怎么忍心下手呢！"
		elseif rst("protect")>tt then
			cast="<font color=FF0000>【投掷】</font>%%现在处于练功保护阶段，所以##无法对之使用"&mg
		else
			stespecial=rst("特技")
			stmoral=rst("道德")
			condition=rst("状态")
			sthp=rst("体力")
			if isnull(atespecial) then atespecial="无"
			conn.execute "update 用户 set 道德=道德-1 where 姓名='"&un&"'"
			conn.execute "update 物品 set 数量=数量-1 where 名称='"&mg&"' and 所有者='"&un&"'"
			if instr(stespecial,";抗"&atespecial&";")<>0 then atespecial="无" 
			if condition="死亡" then
				atespecial="无"
				conn.execute "update 用户 set 道德=道德-1 where 姓名='"&un&"'"
				cast="<font color=FF0000>【投掷】</font>%%业已死亡，##居然还向之投掷"&mg&"，道德因此减1<bgsound src='../mid/anqi.wav' loop=1>"
			elseif sthp>athp then
				conn.execute "update 用户 set 体力=体力-"&athp&",防御=防御+1 where 姓名='"&st&"'"
				especialtime=athp
				cast="<font color=FF0000>【投掷】</font>##向%%投掷了"&mg&"，使之体力减少<bgsound src='../mid/anqi.wav' loop=1>"&athp
			else
				morcal=clng(fmorcal-tmorcal*0.01)
				if morcal>2100000000 then
					morcal=2100000000
				elseif morcal<-2100000000 then
					morcal=-2100000000
				end if	
				atespecial="死亡"
				especialtime=1800
				conn.execute "update 用户 set 道德="&morcal&" where 姓名='"&un&"'"
				conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
				cast="<font color=FF0000>【投掷】</font>##向%%投掷了"&mg&"，并直接造成%%死亡"
			end if
			if atespecial<>"无" then cast=cast&newaberrant(st,un,atespecial,especialtime)
			Response.Write "<script language=javascript>parent.resfrm.location.href="&chr(34)&"miniweapon.asp"&chr(34)&";</script>" 
		end if	
	end if	
	rst.Close
end function
function sendmoney(un,st,mg)
	if not isnumeric(mg) then
		sendmoney="<font color=FF0000>【送钱】</font>##想你的确是送钱"&mg&"给%%吗？"
	else
		mg=clng(mg)	
		if mg<1 then
			sendmoney="<font color=FF0000>【送钱】</font>##想你的确是送钱"&mg&"给%%吗？"
		else	
			rst.Open "select * from 用户 where 姓名='"&un&"' and 银两>="&mg,conn
			if rst.EOF or rst.BOF then
				sendmoney="<font color=FF0000>【送钱】</font>%%笑着对##说：'你的心情我理解，等你有了"&mg&"两银子再来送我好吗？'"
			else
			conn.execute "update 用户 set 银两=银两+"&mg&" where 姓名='"&st&"'"
			conn.execute "update 用户 set 银两=银两-"&mg&" where 姓名='"&un&"'"
			sendmoney="<font color=FF0000>【送钱】</font>##将"&mg&"两银子送给了%%<bgsound src='../mid/thanks.wav' loop=1>"
			end if
		end if	
	end if		
end function
function alert(un,st,co,gr,sn,mg)
	if (co="官府" and gr>=Application("Ba_jxqy_kickguestright")) or (Application("Ba_jxqy_systemname"&sn)=co and gr>7) then
		alert="<font color=FF0000 size=3>【警告】%%,网管慎重警告你:"&mg&"</font>"
	else
		alert="<font color=FF0000>【警告】</font>%%,##向你发无效警告:"&mg
	end if
end function
function bulletin(un,st,co,gr,sn,mg)
	if (co="官府" and gr>=Application("Ba_jxqy_kickguestright")) or (Application("Ba_jxqy_systemname"&sn)=co and gr>7) then
		bulletin="<font color=FF0000 size=3>【公告】</font><table bgcolor=FFFF00 width=80% align=center border=3><tr><td bgcolor=00FF00 align=center>官 府 公 告</td></td><tr><td>"&msg&"</td></tr></table>"
	else
		bulletin="<font color=FF0000>【公告】</font>##，你无权发告公告:"&mg
	end if
end function
function distributemoney(un,co,gr,sn)
	randomize()
	if co="官府" and gr>=Application("Ba_jxqy_unlockipright") then
		money=clng(rnd()*30000)+1000
		onlinelist=Application("Ba_jxqy_onlinelist")
		for i=1 to ubound(onlinelist) step 8
			conn.execute "update 用户 set 银两=银两+"&money&" where 姓名='"&onlinelist(i)&"'"
		next
		erase onlinelist
		distributemoney="<font color=FF0000>【发钱】</font>##代表官府开仓赈灾，每人得到银两<bgsound src='../mid/faqian.wav' loop=1>"&money
	else
		onlinenum=Application("Ba_jxqy_allonlinenum")
		everymoney=clng(rnd()*300)+200
		money=onlinenum*everymoney
		rst.Open "select * from 用户 where 姓名='"&un&"' and 银两>="&money,conn
		if rst.EOF or rst.BOF then
			distributemoney="<font color=FF0000>【发钱】</font>##，你的钱好象不是你想象的那样多哟，下次吧，下次再来请客好了。"
		elseif onlinenum<10 then
			distributemoney="<font color=FF0000>【发钱】</font>##，这儿压根没有几个人，请客没有任何意义，下次吧，下次再来请客好了。"	
		else
			conn.execute "update 用户 set 银两=银两-"&money&",道德=道德+"&onlinenum&" where 姓名='"&un&"'"
			onlinelist=Application("Ba_jxqy_onlinelist")
			for i=1 to ubound(onlinelist) step 8
				conn.execute "update 用户 set 银两=银两+"&everymoney&" where 姓名='"&onlinelist(i)&"'"
			next
			erase onlinelist
			distributemoney="<font color=FF0000>【发钱】</font>##今天心情大好拿出"&money&"两钱子大宴宾朋，在场每人得到"&everymoney&"两银子的利市<bgsound src='../mid/faqian.wav' loop=1>"
		end if
		rst.Close 	
	end if
end function
function degree(un,st,co,mg)
	if mg="掌门" then
		degree="<font color=FF0000>【册封】</font>##,你无法册封%%为掌门"
	else
		rst.Open "select * from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
		if rst.EOF or rst.BOF then
			degree="<font color=FF0000>【册封】</font>##,你不是本派掌门，无法册封%%"
		else
			rst.Close
			rst.Open "select * from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
			if rst.EOF or rst.BOF then
				degree="<font color=FF0000>【册封】</font>%%不是本派弟子，所以##无法册封！"
			else
				conn.execute "update 用户 set 身份='"&mg&"' where  姓名='"&st&"'"
				degree="<font color=FF0000>【册封】</font>##册封%%为"&co&mg&"！"
			end if
		end if
		rst.Close
	end if		
end function
function dislodge(un,st,co)
	rst.Open "select * from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
	if rst.EOF or rst.BOF then
		dislodge="<font color=FF0000>【逐出】</font>##,你不是本派掌门，无法将%%逐出山门"
	else
		rst.Close
		rst.Open "select * from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
		if rst.EOF or rst.BOF then
			dislodge="<font color=FF0000>【逐出】</font>%%不是本派弟子，所以##你无权处置！"
		else
			conn.execute "update 用户 set 身份='无',门派='无' where  姓名='"&st&"'"
			dislodge="<font color=FF0000>【逐出】</font>##将%%无情地逐出了"&co&"！"
		end if
	end if
	rst.Close
end function
function impart(un,st,mg)
	if not isnumeric(mg) then
		impart="<font color=FF0000>【传功】</font>##想你的确是传功"&mg&"给%%吗？"
	else
		mg=clng(mg)	
		if mg<100 then
			impart="<font color=FF0000>【传功】</font>##想你的确是传功"&mg&"给%%吗？也太少了点儿吧"
		else	
			rst.Open "select * from 用户 where 姓名='"&un&"' and 内力>="&mg,conn
			if rst.EOF or rst.BOF then
				impart="<font color=FF0000>【传功】</font>%%笑着对##说：'其实你只传给我几百我也会很开心的不用那么"&mg&"多！'"
			else
				conn.execute "update 用户 set 内力=内力+"&mg*0.8&",道德=道德+"&mg*0.01&" where 姓名='"&st&"'"
				conn.execute "update 用户 set 内力=内力-"&mg&" where 姓名='"&un&"'"
				impart="<font color=FF0000>【传功】</font>##将"&clng(mg*0.8)&"的内力传授给了%%<bgsound src='../mid/xixing.wav' loop=1>"
			end if
		end if	
	end if
end function
function burglemp(un,st)
	randomize()
	odds=rnd()*10000
	if odds>1999 then
		rst.Open "select * from 用户 where 姓名='"&un&"'",conn
		uhp=rst("体力")	
		rst.Close
		if odds*0.2>uhp then odds=uhp
		conn.execute "update 用户 set 体力=体力-"&odds*0.2&",道德=道德-"&odds*0.2&" where 姓名='"&un&"'"
		burglemp="<font color=FF0000>【吸星大法】</font>##因在公众场合对%%滥施吸星大法，有违江湖公道，被大家打得动弹不得！体力和道德下降"&clng(odds*0.2)&"，"&newaberrant(un,"大家","麻痹",120)
	else
		rst.Open "select * from 用户 where 姓名='"&st&"'",conn
		smp=rst("内力")
		rst.Close
		if odds>smp then odds=smp
		if odds>1900 then odds=smp*0.5
		conn.execute "update 用户 set 内力=内力+"&odds&",道德=道德-"&odds*0.1&" where 姓名='"&un&"'"
		conn.execute "update 用户 set 内力=内力-"&odds&" where 姓名='"&st&"'"
		burglemp="<font color=FF0000>【吸星大法】</font>##偷偷使用了江湖中最忌讳的吸星大法，吸走了%%"&clng(odds)&"的内力<bgsound src='../mid/xixing.wav' loop=1>"
	end if	
end function
function burglemoney(un,st)
	randomize()
	odds=rnd()*10000
	rst.Open "select * from 用户 where 姓名='"&un&"'",conn
	uhp=rst("体力")	
	rst.Close
	if odds>3999 then
		if odds*0.2>uhp then odds=uhp
		conn.execute "update 用户 set 体力=体力-"&odds*0.2&",道德=道德-"&odds*0.2&" where 姓名='"&un&"'"
		burglemoney="<font color=FF0000>【偷窃】</font>##因在公众场合对%%实施武力抢劫，有违江湖公道，被大家打疯了，体力和道德下降"&clng(odds*0.2)&"，"&newaberrant(un,"大家","疯狂",180)
	elseif odds>0 and odds<4000 then
		rst.Open "select * from 物品  where 所有者='"&st&"' and 数量>0 order by 价格 desc",conn
		if rst.EOF or rst.BOF then
			odds=clng(odds*0.2)
			if odds>uhp then odds=uhp
			conn.execute "update 用户 set 体力=体力-"&odds&",道德=道德-"&odds&" where 姓名='"&un&"'"
			burglemoney="<font color=FF0000>【偷窃】</font>%%形容落魄，身无长物，##居然想对之进行偷窃，被大家打的遍体鳞伤，体力和道德下降"&odds&"，"&newaberrant(un,"大家","封咒",180)
		else
			cname=rst("名称")
			set rst2=server.CreateObject("adodb.recordset")
			rst2.Open "select * from 物品 where 所有者='"&un&"' and 名称='"&cname&"'",conn
			if rst2.EOF or rst2.BOF then
				conn.execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&cname&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")&",'"&un&"',false,false)"
			else
				conn.execute "update 物品 set 数量=数量+1 where 所有者='"&un&"' and 名称='"&cname&"'"
			end if
			conn.execute "update 物品 set 数量=数量-1 where 所有者='"&st&"' and 名称='"&cname&"'"
			rst2.Close
			set rst2=nothing
			burglemoney="<font color=FF0000>【偷窃】</font>##偷偷使出妙手空空的手段，偷走了%%<bgsound src='../mid/xiaotou.wav' loop=1>随身携带的"&cname
		end if
		rst.Close
	else
		rst.Open "select * from 用户 where 姓名='"&st&"'",conn
		umoney=rst("银两")	
		rst.Close
		if odds>umoney then odds=umoney
		if odds>1900 then odds=umoney*0.5
		conn.execute "update 用户 set 银两=银两+"&odds&",道德=道德-"&odds*0.01&" where 姓名='"&un&"'"
		conn.execute "update 用户 set 银两=银两-"&odds&" where 姓名='"&st&"'"
		burglemoney="<font color=FF0000>【偷窃】</font>##偷偷使出妙手空空的手段，偷走了%%"&clng(odds)&"的银两<bgsound src='../mid/xiaotou.wav' loop=1>"
	end if	
end function
function present(un,st,mg)
	rst.Open "select * from 物品  where 所有者='"&un&"' and 数量>0 and 名称='"&mg&"'",conn
	if rst.EOF or rst.BOF then
		present="<font color=FF0000>【赠送】</font>##形容落魄，身无长物，居然也想拿"&mg&"送给%%。"
	else
		set rst2=server.CreateObject("adodb.recordset")
		rst2.Open "select * from 物品 where 所有者='"&st&"' and 名称='"&mg&"'",conn
		if rst2.EOF or rst2.BOF then
			conn.execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&mg&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")&",'"&st&"',false,false)"
		else
			conn.execute "update 物品 set 数量=数量+1 where 所有者='"&st&"' and 名称='"&mg&"'"
		end if
		conn.execute "update 物品 set 数量=数量-1 where 所有者='"&un&"' and 名称='"&mg&"'"
		rst2.Close
		set rst2=nothing
		present="<font color=FF0000>【赠送】</font>##将"&mg&"大大方方的递到%%面前说：“一点心意，菲薄难堪！”"
	end if
	rst.Close
end function
function card(un,st,mg)
	if Application("Ba_jxqy_fellow")=True then
		rst.Open "select tc.* from card tc inner join 用户 tu ON tu.姓名=tc.username where tc.username='"&un&"' and tc.cname='"&mg&"' and tc.cnumber>0 and tu.会员=true",conn
	else
		rst.Open "select * from card where username='"&un&"' and cname='"&mg&"' and cnumber>0",conn
	end if	
	if rst.EOF or rst.BOF then
		card="<font color=FF0000>【道具】</font>##，你并没有"&mg&"可对%%使用！"
	else
		especial=trim(rst("cespecial"))
		especialtime=rst("ctime")
		esptime=dateadd("s",especialtime,now())
		rst.Close
		rst.Open "select 特技 from 用户 where 姓名='"&st&"'"
		if rst.EOF or rst.bof then Response.End 
		if instr(rst("特技"),";抗"&especial&";")<>0 then
			card="<font color=FF0000>【道具】</font>##对%%使用了"&mg&",但是%%具有抗"&especial&"特性，因而没有任何效果"
		else
			card="<font color=FF0000>【道具】</font>##对%%使用了"&mg&newaberrant(st,un,especial,especialtime)
		end if
		conn.execute "update card set cnumber=cnumber-1 where username='"&username&"' and cname='"&mg&"'"
		if instr(";死亡;入狱;逮捕;",especial)<>0 then conn.execute "update 用户 set 状态='"&especial&"',最后登录时间='"&esptime&"' where 姓名='"&st&"'"
	end if
	rst.Close 
end function
function newaberrant(st,un,especial,especialtime)
	nowtime=now()
	aberrantlist=Application("Ba_jxqy_aberrantlist")
	aberrantlistubd=ubound(aberrantlist)
	dim newaberrantlist()
	newaberrantname=";"
	j=1
	addsucf=false
	for i=1 to aberrantlistubd step 4
		if addsucf=false and aberrantlist(i+2)=especial and aberrantlist(i)=st then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=st
			newaberrantlist(j+1)=un
			newaberrantlist(j+2)=especial
			newaberrantlist(j+3)=dateadd("s",especialtime,aberrantlist(i+3))
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
			addsucf=true
			if especial<>"死亡" then newaberrant=newaberrant&","&especial&"程度加深"&especialtime
		elseif datediff("s",aberrantlist(i+3),timetmp)<0 then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=aberrantlist(i)
			newaberrantlist(j+1)=aberrantlist(i+1)
			newaberrantlist(j+2)=aberrantlist(i+2)
			newaberrantlist(j+3)=aberrantlist(i+3)
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
		end if
	next
	if addsucf=false then
		redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
		newaberrantlist(j)=st
		newaberrantlist(j+1)=un
		newaberrantlist(j+2)=especial
		newaberrantlist(j+3)=dateadd("s",especialtime,nowtime)
		newaberrantname=newaberrantname&st&";"
		if especial<>"死亡" then newaberrant=newaberrant&","&especial&especialtime&"秒"
	end if
	Application.Lock
	Application("Ba_jxqy_aberrantlist")=newaberrantlist
	Application("Ba_jxqy_aberrantname")=newaberrantname
	Application.UnLock
	erase newaberrantlist
	erase aberrantlist
	if especial="麻痹" then
		newaberrant=newaberrant&"<bgsound src='../mid/dianxu.wav' loop=1>"
	elseif especial="死亡" then
		lastlogintime=dateadd("s",especialtime,nowtime)
		lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
		conn.execute "update 用户 set 状态='死亡',最后登录时间="&lastlogintimetype&" where 姓名='"&st&"'"
		newaberrant=newaberrant&"<bgsound src='../mid/oh_no.wav' loop=1>"
	end if
end function
%>
