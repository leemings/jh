<SCRIPT LANGUAGE="JavaScript">if(window.name!="d"){var i=1;while(i<=50){window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"}</script>
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
	Response.Write "<script language=JavaScript>{alert('又没有人想亲你，臭美！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('提示：你想亲的人已经掉出聊天室或已经走了！');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("aqjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"不想亲你！');}</script>"
	Response.End
end if
if fromname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('发神经吗,哪有自已亲自已的有病！');}</script>"
	Response.End
end if
if qql<>"飞吻" and qql<>"亲亲" and qql<>"狂吻" and qql<>"拒绝" then
	Response.Write "<script language=JavaScript>{alert('江湖暂时无此飞吻！"&QQ1&"');}</script>"
	Response.End
end if
nowsj=DateDiff("s",sj,now())
if nowsj>30 then
	Application.Lock
	Application("aqjh_qw")=""
	Application.unLock
	Response.Write "<script language=JavaScript>{alert('提示：提示：超过30秒未接受亲吻请求！');}</script>"
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
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的道德没达到1000或魅力不到5000，和他亲吻太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<10000 or rs("内力")<6000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的体力不足10000或内力不足6000，小心别把他累着了，不同亲吻耗力不同！');}</script>"
	Response.End
end if
toml=rs("魅力")
fromsex=rs("性别")
rs.close
rs.open "select 体力,内力,道德,魅力 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人！');}</script>"
	Response.End
end if
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_qw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的道德没达到1000或魅力不到5000，和你亲吻太没面子了！');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('你的体力不足10000或内力不足6000，小心虚脱！不同亲吻耗力不同！');}</script>"
	Response.End
end if
myml=rs("魅力")
rs.close
'开始亲吻
ab="<font color=red><b>【抛个飞吻】</b></font>"
select case qql
	case "飞吻"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp51.gif>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>大眼睛"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>含情的和<font color=blue>"& fromname &"</font>对视"& tp &"，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>50点</font>，魅力上涨<font color=red>50点</font>，"& fromname &"魅力上涨100点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“我好喜欢你呀……”"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“你真好……”"& tp &"<font color=blue>"& fromname &"</font>陶醉梦幻之中。"
					end if
				else
					if fromsex="女" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“是真的吗……”"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“你真好……”"& tp &"<font color=blue>"& fromname &"</font>陶醉在梦幻之中。"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>50点</font>，魅力上涨<font color=red>50点</font>，"& fromname &"魅力上涨100点。"
			case 3
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>大眼睛"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>深情的和<font color=blue>"& fromname &"</font>对视"& tp &"，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>50点</font>，魅力上涨<font color=red>50点</font>，"& fromname &"魅力上涨100点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+50,魅力=魅力+50,体力=体力-4000,内力=内力-2500 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+100,体力=体力-4000,内力=内力-2500 where 姓名='"&fromname&"'"
case "亲亲"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=pic/dp78.GIF>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>的小嘴儿"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>深情的和<font color=blue>"& fromname &"</font>对视"& tp &"，两人陶醉在梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“好喜欢你呀……”，"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“你真好……”，"& tp &"<font color=blue>"& fromname &"</font>陶醉在梦幻之中。"
					end if
				else
					if fromsex="女" then
						x="<font color=blue>"& fromname &"</font>站在那里小声说着：“你真好……”，"& tp &"<font color=blue>"& aqjh_name &"</font>陶醉在浪漫之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>站在那里小声说着：“是真的吗……”,"& tp &"<font color=blue>"& fromname &"</font>陶醉在浪漫之中。"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
			case 3
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>深情的看着<font color=blue>"& aqjh_name &"</font>的小嘴儿"& tp &"，两人陶醉在浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>深情的和<font color=blue>"& fromname &"</font>对视"& tp &"，两人陶醉梦幻之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>80点</font>，魅力上涨<font color=red>80点</font>，"& fromname &"魅力上涨150点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+80,魅力=魅力+80,体力=体力-6500,内力=内力-4000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+150,体力=体力-6500,内力=内力-4000 where 姓名='"&fromname&"'"
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
		conn.execute "update 用户 set 道德=道德+120,魅力=魅力+120,体力=体力-10000,内力=内力-6000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+200,体力=体力-10000,内力=内力-6000 where 姓名='"&fromname&"'"
	
	case "拒绝"
		randomize timer
		r=int(rnd*5)/5
		toml=int(toml*r)
		ab="<font color=red><b>【拒绝亲吻】</b></font><font color=blue>"& aqjh_name &"</font>拒绝了<font color=blue>"& fromname &"</font>的亲吻邀请，不认识你干吗要亲，你找4啦……<img src=pic/dp33.gif>说话间举起左手打了过去。被巴掌打的晕晕的说道：白天为什么看见了星星？好晕。"& fromname &"碰了一鼻子灰，"& tp &"<font color=blue>"& fromname &"</font>。魅力顿时下降500。"& aqjh_name &"魅力上升100点，道德下降50点。"
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