<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">if(window.name!="d"){var i=1;while(i<=50){window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"}</script>
<!--#include file="../const.asp"-->
<%
qql=lcase(trim(request.querystring("zl"))) '亲密接触种类
if qql="" then
	Response.Write "<script language=javascript>{alert('数据出错！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if application("aqjh_tw")="" then
	Response.Write "<script language=JavaScript>{alert('又没有人想亲你！');}</script>"
	Response.End
end if
wddata=split(application("aqjh_tw"),"|")
fromname=wddata(0)
toname=wddata(1)
sj=wddata(2)
erase wddata

if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：想亲你的人已经掉出聊天室或已经走了！');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("aqjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>aqjh_name then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"又不是在亲你，少在这自作了！');}</script>"
	Response.End
end if
if fromname=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('发神经吗,那有自已亲自已的！');}</script>"
	Response.End
end if
if qql<>"去酒店" and qql<>"拒绝" and qql<>"进房间" and qql<>"在江湖大厅" then
	Response.Write "<script language=JavaScript>{alert('江湖暂时只有去酒店，进房间，在江湖的大厅和拒绝功能！其它功能还在开发中"&QQ1&"');}</script>"
	Response.End
end if
nowsj=DateDiff("s",sj,now())
if nowsj>50 then
	Application.Lock
	Application("aqjh_tw")=""
	Application.unLock
	Response.Write "<script language=JavaScript>{alert('提示：超过50秒未接受亲密接触！');}</script>"
	Response.End
end if
Application.Lock
	Application("aqjh_tw")=""
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
if rs("道德")<1 or rs("魅力")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的道德没达到100或魅力不到100，色他（她）太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<100 or rs("内力")<100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的体力不足100或内力不足100，小心别把他累着了！');}</script>"
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
if rs("道德")<100 or rs("魅力")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的道德没达到100或魅力不到500，给你色太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<800 or rs("内力")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("aqjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的体力不足800或内力不足500，小心虚脱！');}</script>"
	Response.End
end if
myml=rs("魅力")
rs.close
'开始求签
ab="<font color=red><b>【亲密接触成功】【魅力四射】</b></font><font color=blue><b>"& aqjh_name &"</b></font>答应了<font color=blue><b>"& fromname &"</b></font>的亲密接触。<br>"
select case qql
	case "去酒店"
		randomize timer
		r=int(rnd*2)+1
		tp="<img src=img/11.jpg>"
		select case r
			case 1
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>拉起<font color=blue>"& aqjh_name &"</font>开着自己心爱的小车来到了“中国大酒店”"& tp &"，两人陶醉在这夜的浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>拉起<font color=blue>"& fromname &"</font>开着自己心爱的小车来到了“中国大酒店”"& tp &"，两人陶醉在这夜的浪漫之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
			case 2
				randomize timer
				m=int(rnd*2)+1
				if m=1 or m=3 then
					if fromsex="男" then
						x="<font color=blue>"& fromname &"</font>和<font color=blue>"& aqjh_name &"</font>在酒店的房间里窃窃私语。吵得隔壁房间正在偷情的情人睡不着。"& tp &"<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>却好似旁若无人般地沉浸在优美的旋律之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>在酒店的房间里窃窃私语。吵得隔壁房间正在偷情的情人睡不着。"& tp &"<font color=blue>"& fromname &"</font>和<font color=blue>"& aqjh_name &"</font>却好似旁若无人般地沉浸在优美的旋律之中。"
					end if
				else
					if fromsex="女" then
						x="<font color=blue>"& fromname &"</font>和<font color=blue>"& aqjh_name &"</font>在酒店的房间里窃窃私语。吵得隔壁房间正在偷情的情人睡不着。"& tp &"<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>却好似旁若无人般地沉浸在优美的旋律之中。"
					else
						x="<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>在酒店的房间里窃窃私语。吵得隔壁房间正在偷情的情人睡不着。"& tp &"<font color=blue>"& fromname &"</font>和<font color=blue>"& aqjh_name &"</font>却好似旁若无人般地沉浸在优美的旋律之中。"
					end if
				end if
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
			case 3
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>拉起<font color=blue>"& aqjh_name &"</font>开着自己心爱的小车来到了“中国大酒店”"& tp &"，两人陶醉在这夜的浪漫之中。"
				else
					x="<font color=blue>"& aqjh_name &"</font>拉起<font color=blue>"& fromname &"</font>开着自己心爱的小车来到了“中国大酒店”"& tp &"，两人陶醉在这夜的浪漫之中。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+100,魅力=魅力+100,体力=体力-8000,内力=内力-5000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+500,体力=体力-8000,内力=内力-5000 where 姓名='"&fromname&"'"
	case "进房间"
		randomize timer
		r=int(rnd*2)+1
		r=1
		tp="<img src=img/jgj.gif>"
'		select case r
'			case 1
				if fromsex="男" then
					x="<font color=blue>"& fromname &"</font>轻轻地搂着<font color=blue>"& aqjh_name &"</font>，来到两人共建的小屋"& tp &"，浪漫的拥抱！"
				else
					x="<font color=blue>"& aqjh_name &"</font>轻轻地搂着<font color=blue>"& fromname &"</font>，来到两人共建的小屋"& tp &"，浪漫的拥抱！"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
'		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+100,魅力=魅力+100,体力=体力-8000,内力=内力-5000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+500,体力=体力-8000,内力=内力-5000 where 姓名='"&fromname&"'"
	case "在江湖大厅"
		randomize timer
		r=int(rnd*2)+1
		sjt=int(rnd*6)+1
		select case sjt
			case 1
				tp="<img src=img/f6.gif><img src=img/f2.gif>"
			case 2
				tp="<img src=img/dp68.gif><img src=img/dp63.gif><img src=img/dp61.gif>"
			case 3
				tp="<img src=img/dp78.gif>"
			case 4
				tp="<img src=img/rl.gif><img src=img/f6.gif>"
			case 5
				tp="<img src=img/dp16.gif>"
			case 6
				tp="<img src=img/dp18.gif>"
			case else
				tp="<img src=img/f3.gif>"
		end select
		select case r
			case 1
				if fromsex="男" then
					x=tp&"<font color=blue>"& aqjh_name &"</font>在江湖的大厅里疯狂的拥抱，还不时地向大家抛几个飞吻，弄的聊天室里怪叫连连，<font color=blue>"& fromname &"</font>也疯狂地扭动着那肥大的臀部。"
				else
					x=tp&"<font color=blue>"& fromname &"</font>在江湖的大厅里疯狂的拥抱，还不时地向大家抛几个飞吻，弄的聊天室里怪叫连连，<font color=blue>"& aqjh_name &"</font>也疯狂地扭动着那肥大的臀部。"
				end if
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
			case 2
				x=tp&"<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>躲在无人的角落里~~~~~~~~~~~~~真是羡慕死人了！！"
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
			case 3
				x=tp&"<font color=blue>"& aqjh_name &"</font>和<font color=blue>"& fromname &"</font>躲在无人的角落里~~~~~~~~~~~~~真是羡慕死人了！！"
				x=x&aqjh_name&"道德上涨<font color=red>100点</font>，魅力上涨<font color=red>100点</font>，"& fromname &"魅力上涨500点。"
		end select
		ab=ab&x
		conn.execute "update 用户 set 道德=道德+100,魅力=魅力+100,体力=体力-8000,内力=内力-5000 where 姓名='"&aqjh_name&"'"
		conn.execute "update 用户 set 魅力=魅力+500,体力=体力-8000,内力=内力-5000 where 姓名='"&fromname&"'"
	case "拒绝"
		randomize timer
		r=int(rnd*5)/10
		toml=int(toml*(1-r))
		ab="<font color=red><b>【拒绝亲密接触】</b></font><font color=blue>"& aqjh_name &"</font>拒绝了<font color=blue>"& fromname &"</font>的亲密接触,就凭你这德性也想来色我？"& fromname &"碰了一鼻子灰，魅力顿时下降"& r*100 &"％。"& aqjh_name &"魅力上升300点，道德下降100点。"
		conn.execute "update 用户 set 魅力="&toml&" where 姓名='"&fromname&"'"
		conn.execute "update 用户 set 魅力=魅力+300,道德=道德-100 where 姓名='"&aqjh_name&"'"
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
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway & ","& nowinroom & ");<"&"/script>"
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