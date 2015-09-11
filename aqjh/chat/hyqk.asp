<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if application("aqjh_qw")="" then
	Response.Write "<script language=JavaScript>{alert('又没有人想请你，臭美！');}</script>"
	Response.End
end if
wenndata=split(application("aqjh_qw"),"|")
fromname=wenndata(0)
toname=wenndata(1)
sj=wenndata(2)
erase wenndata
qql=lcase(trim(request.querystring("zl"))) '亲吻种类
if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：你想请客的人已经掉出聊天室或已经走了！');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("aqjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"没有请你！');}</script>"
	Response.End
end if
if fromname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('发神经吗,哪有自已请自已的有病！');}</script>"
	Response.End
end if
if qql<>"吃饭" and qql<>"冰淇淋" and qql<>"拒绝" then
	Response.Write "<script language=JavaScript>{alert('江湖暂时无此请客！"&QQ1&"');}</script>"
	Response.End
end if
nowsj=DateDiff("s",sj,now())
if nowsj>30 then
	Application.Lock
	Application("aqjh_qw")=""
	Application.unLock
	Response.Write "<script language=JavaScript>{alert('提示：提示：超过30秒未接受请客请求！');}</script>"
	Response.End
end if
Application.Lock
	Application("aqjh_qw")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 体力,内力,道德,魅力,性别 from 用户 where 姓名='"&fromname&"'",conn,2,2
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('江湖中没有"&fromname&"这个人！');}</script>"
	Response.End
end if
if rs("道德")<10 or rs("魅力")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的道德没达到10或魅力不到50，和他一起太没面子了！');}</script>"
	Response.End
end if
toml=rs("魅力")
fromsex=rs("性别")
rs.close
rs.open "select 体力,内力,道德,魅力,武功 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人！');}</script>"
	Response.End
end if
if rs("道德")<10 or rs("魅力")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_qw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的道德没达到10或魅力不到00，和你一起太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<10000 or rs("内力")<6000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_qw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的体力不足10000或内力不足6000，小心死死！不同类型耗力不同！');}</script>"
	Response.End
end if
myml=rs("魅力")
rs.close
'开始亲吻
ab="<font color=red><b>【掏钱请客】</b></font>"
select case qql
	case "吃饭"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=img/chifan.gif>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=blue><font color=blue>"& fromname &"</font>满足的看着<font color=blue>"& aqjh_name &"</font>大口大口地<img src=img/cf1.gif>，"& fromname &"陶醉在"& aqjh_name &"的吃饭气氛之中。"
				else
					x="<font color=blue><font color=blue>"& aqjh_name &"</font>开心的和<font color=blue>"& fromname &"</font>共享<img src=img/cf4.gif>，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>50点</font>，魅力上涨<font color=red>50点</font>，"& fromname &"魅力上涨100点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=blue>"& fromname &"</font>掏出大把钱道：<font color=blue>“走~~今天请你吃顿好的……”<font color=blue>"& aqjh_name &"</font>高兴的在五星级酒店点了一客<img src=img/cf5.gif>，好不客气的吃起来。"
					else
						x="<font color=blue>"& aqjh_name &"</font>对<font color=blue>"& fromname &"</font>说道：<font color=blue>“请我吃饭好吗，我实在太饿了……”<font color=blue>"& fromname &"</font>不耐烦的在路边买了一盘<img src=img/cf6.gif>请客。"
					end if
				else
					if fromsex="女" then
						x="<font color=blue>"& fromname &"</font>对<font color=blue>"& aqjh_name &"</font>说：<font color=blue>嘿嘿~~！走啦~~干嘛那么见外，别跟我客气，就那几个钱算什么，慢慢<img src=img/cf3.gif>，吃后还有！"& fromname &"大口大口的吃了"& aqjh_name &"请的客并说道：真有你的，以后你就是我兄弟了~！！"
					else
						x="<font color=blue>"& fromname &"</font>对<font color=blue>"& aqjh_name &"</font>说：<font color=blue>好哥们，饿了吧！请你到大酒楼吃大菜，别跟我客气，出门在外嘛，慢慢<img src=img/cf2.gif>（大闸蟹真不是盖的，味道好极了！），吃后再聊！"& aqjh_name &"大口大口的吃了"& fromname &"请的客，，临走时抛下句：好哥儿~！！"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>50点</font>，魅力上涨<font color=red>50点</font>，"& fromname &"魅力上涨100点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 银两=银两-20000,道德=道德+50,魅力=魅力+50 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 道德=道德+50,魅力=魅力+100 where 姓名='"&fromname&"'"
case "冰淇淋"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp78.GIF>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=0088FF><font color=blue>"& fromname &"</font>揽着兄弟<font color=blue>"& aqjh_name &"</font>的肩膀，看着炎热的天气，<font color=blue>"& aqjh_name &"</font>提议说：我请你吃冰淇淋吧，我知道一家很好吃的哦~说罢二人走进冰心冰淇淋店，点了2杯<img src=img/bql1.gif>，<font color=blue>"& aqjh_name &"</font>看着<font color=blue>"& fromname &"</font>陶醉在冰淇淋的世界中。"
				else
					x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>恶狠狠的和<font color=blue>"& fromname &"</font>对视：走~！哥们~！~这么炎热的天气~请你吃冰淇淋去，你看<img src=img/bql2.gif>，好漂亮，好好吃，放心吃吧，我付钱~~~！~"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=0088FF><font color=blue>"& fromname &"</font>站在那里小声嘀咕说着：“你想吃就跟我说嘛……”，<font color=blue>"& fromname &"</font>从冰心冰淇淋店拿出一大盘<img src=img/bql3.gif>，请<font color=blue>"& aqjh_name &"</font>吃多多的。"
					else
						x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>大喝道：“你小子是不是看不起我啊~！请你吃个冰淇淋你推三落四的，今天你不吃下他我跟你没完~！！！真气人~！”，<font color=blue>"& fromname &"</font>看着<font color=blue>"& aqjh_name &"</font>的脸，吃着<font color=blue>"& fromname &"</font>请吃的<img src=img/bql4.gif>。"
					end if
				else
					if fromsex="女" then
						x="<font color=0088FF><font color=blue>"& fromname &"</font>泪汪汪得看着<font color=blue>"& aqjh_name &"</font>说：“我只剩下这一杯冰淇淋了，你想要就给你……”，<font color=blue>"& aqjh_name &"</font>二话不说，从<font color=blue>"& fromname &"</font>手里抢过<img src=img/bql5.gif>。只把<font color=blue>"& aqjh_name &"</font>气得直跺脚。"
					else
						x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>站在冰心大马路旁，嘴里嘀咕着：“想吃冰淇淋，我要吃……”,看到<font color=blue>"& fromname &"</font>正买好<img src=img/bql6.gif>在掏钱付，一把跑过马路，不料<font color=blue>"& fromname &"</font>瞧见后明白有人要抢，不经意大喊一声，<font color=blue>"& aqjh_name &"</font>吓得溜回老家，在半路含恨而终…………"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
			case 3
				if fromsex="男" then
					x="<font color=0088FF><font color=blue>"& fromname &"</font>站在冰淇淋店旁，吃着刚买的<img src=img/bql3.gif>，手里另拿着一杯（这人真贪吃）看着<font color=blue>"& aqjh_name &"</font>的小嘴儿，发现………<font color=blue>"& aqjh_name &"</font>在留口水……<font color=blue>"& fromname &"</font>猛把手里的冰淇淋晃来晃去，只把<font color=blue>"& aqjh_name &"</font>气得晕过去了。"
				else
					x="<font color=0088FF><font color=blue>"& aqjh_name &"</font>吃着<font color=blue>"& fromname &"</font>请的<img src=img/bql2.gif>，越吃越有劲，叫<font color=blue>"& fromname &"</font>再请一杯，<font color=blue>"& fromname &"</font>把手往钱袋一拿，只剩1毛钱，<font color=blue>"& aqjh_name &"</font>大声说道：穷鬼~~~~~~~~。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+80,魅力=魅力+80,体力=体力-650,内力=内力-400 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+150,体力=体力-6500,内力=内力-400 where 姓名='"&fromname&"'"
case "狂吻"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp76.gif>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>大眼睛"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>深情的看着<font color=blue>"& fromname &"</font>小嘴儿"& tp &"，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>120点</font>，魅力上涨<font color=red>120点</font>，"& fromname &"魅力上涨200点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“好喜欢你呀……”，还是被众人听到。"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“你真好……”，还是被众人听到。"& tp &"<font color=blue>"& fromname &"</font>陶醉在梦幻之中。"
					end if
				else
					if fromsex="女" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“好喜欢你呀……”，还是被众人听到。"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“真的吗……”，还是被众人听到。"& tp &"<font color=blue>"& fromname &"</font>陶醉在梦幻之中。"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>120点</font>，魅力上涨<font color=red>120点</font>，"& fromname &"魅力上涨200点。"
			case 3
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>的小嘴儿"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>深情的看着<font color=blue>"& fromname &"</font>的大眼睛"& tp &"，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>120点</font>，魅力上涨<font color=red>120点</font>，"& fromname &"魅力上涨200点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+120,魅力=魅力+120,体力=体力-1000,内力=内力-6000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+200,体力=体力-1000,内力=内力-600 where 姓名='"&fromname&"'"
	
	case "拒绝"
		randomize timer
		r=int(rnd*5)/5
		toml=int(toml*r)
		ab="<font color=red><b>【拒绝请客】</b></font><font color=blue>"& aqjh_name &"</font>拒绝了<font color=blue>"& fromname &"</font>的请客，你看不起我啊，你找4啦……我要吃我自己没钱么？"& fromname &"碰了一鼻子灰，吃力不讨好。魅力顿时下降500。"& aqjh_name &"魅力上升100点，道德下降50点。"
		conn.execute "update 用户 set 魅力=魅力-500 where 姓名='"&fromname&"'"
		conn.execute "update 用户 set 魅力=魅力+100,道德=道德-50 where 姓名='"&aqjh_name&"'"
end select
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says=ab	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho=fromname
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>