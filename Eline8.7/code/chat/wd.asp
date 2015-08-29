<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if

fromname=LCase(trim(request.querystring("fromname")))	'对方名字
toname=LCase(trim(request.querystring("toname")))	'自已名字
qql=lcase(trim(request.querystring("qq"))) '跳舞种类
if instr(fromname,"or")<>0 or instr(fromname,".")<>0 or instr(fromname,"'")<>0 or instr(fromname,"OR")<>0 or instr(fromname,"%20")<>0 or instr(fromname,">")<>0 or instr(fromname,"=")<>0 or instr(fromname,"<")<>0 or instr(fromname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if instr(toname,"or")<>0 or instr(toname,".")<>0 or instr(toname,"'")<>0 or instr(toname,"OR")<>0 or instr(toname,"%20")<>0 or instr(toname,">")<>0 or instr(toname,"=")<>0 or instr(toname,"<")<>0 or instr(toname," ")<>0 then
	Response.Write "<script language=JavaScript>{alert('你想作什么！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(fromname)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：你的舞伴不在聊天室中，不可以跳舞！');location.href = 'javascript:history.go(-1)';}</script>"
	Application.Lock
	Application("sjjh_smxs")=""
	Application.unLock
	Response.End
end if
if toname<>sjjh_name then
	Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"不是在请你跳舞，少在这自作了！');}</script>"
	Response.End
end if
if fromname=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('发神经吗,那有自已请自已跳舞的！');}</script>"
	Response.End
end if
if qql<>"华尔兹" and qql<>"探戈" and qql<>"迪斯高" then
	Response.Write "<script language=JavaScript>{alert('本处暂时只有华尔兹、探戈和迪斯高三种舞曲！');}</script>"
	Response.End
end if
Application.Lock
	Application("sjjh_tw")=""
Application.unLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的道德没达到1000或魅力不到5000，和他跳舞太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&fromname&" 的体力不足8000或内力不足5000，小心别把他累着了！');}</script>"
	Response.End
end if
toml=rs("魅力")
fromsex=rs("性别")
rs.close
rs.open "select 体力,内力,银两,道德,魅力 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人！');}</script>"
	Response.End
end if
if rs("银两")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你没有5万两银子，那来的门票？');}</script>"
	Response.End
end if
if rs("道德")<1000 or rs("魅力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("sjjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的道德没达到1000或魅力不到5000，和你跳舞太没面子了！');}</script>"
	Response.End
end if
if rs("体力")<8000 or rs("内力")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Application.Lock
	Application("sjjh_tw")=""
	Application.UnLock
	Response.Write "<script language=JavaScript>{alert('你的体力不足8000或内力不足5000，小心虚脱！');}</script>"
	Response.End
end if
myml=rs("魅力")
rs.close
'开始求签
ab="【"&qql&"舞】<font color=red>※"& sjjh_name &"※</font>和<font color=red>※"& fromname &"※</font>跳起了"&qql&"舞。"
wei="<font color=red>※"& sjjh_name &"※</font>"
yinliang=50000
mytl=0
mynl=0
myml=0
myallvalue=0
myjb=0
mydd=0
select case qql
	case "华尔兹"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/hez.gif><font color=blue>"& fromname &"</font>看到<font color=blue>"& sjjh_name &"</font>在侨都酒店舞厅里这样的不自然，就勇敢大方地主动邀他跳上一曲悠扬的华尔兹，两人陶醉在这夜的浪漫之中。加收15万两银子"
				yingliang=yingliang+150000
			case 2,8
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& fromname &"</font>轻轻地走过来，搂着<font color=blue>"& sjjh_name &"</font>，伴随着美妙的华尔兹音乐，%%投射过来异样的眼神，诧异也好，欣赏也罢！并不曾使的舞步凌乱.因为令飞扬的，不是"& sjjh_name &"注视的目光，而是在心目合一，忘我地敞连在一起！"
			case 3,7
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>轻轻地搂着<font color=blue>"& fromname &"</font>，优雅的舞着醉人的华尔兹，在拥挤的人群中，众人投射过来异样的眼神，诧异也好，欣赏也罢！并不曾使的舞步凌乱.因为令飞扬的，不是"& fromname &"注视的目光，而是年轻的心！"
			case 4
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>傻呼呼的走到<font color=blue>"& fromname &"</font>面前，唉，可怜呀，你不会就别找我呀，跳舞就像踏步一样，还踏了"& fromname &"的脚，一起受罚了，扣"& fromname &"十万银两，！"&sjjh_name&"痛失50点卷！"
				dianjuan=-50
				yinliang=-100000
			case 5
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>傻傻地搂着<font color=blue>"& fromname &"</font>，笨笨的舞着醉人的华尔兹，在拥挤的人群中，使得舞步凌乱.扣"& fromname &"十万银两，！"&sjjh_name&"痛失50点卷"
				dianjuan=-50
				yinliang=-100000
			case 6
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>轻轻地搂着<font color=blue>"& fromname &"</font>，优雅的舞着醉人的华尔兹，在拥挤的人群中，使得舞步飞扬.加"& fromname &"十八万银两，！"&sjjh_name&"得50点卷"
				dianjuan=50
				yingliang=180000
			case else
				tw="<bgsound src=mid/hez.mid loop=1><img src=img/hez.gif><font color=blue>"& sjjh_name &"</font>轻轻地搂着<font color=blue>"& fromname &"</font>，优雅的舞着醉人的华尔兹，哈哈，跳舞不找我找谁，真是舞林高脚，"& fromname &"幸福了，有这么好的伴了！"
		end select
	case "探戈"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>拉起<font color=blue>"& sjjh_name &"</font>的小手翩翩跳起了酒醉的探戈，两人陶醉在这夜的浪漫之中。"
			case 2,8
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>乱拉<font color=blue>"& sjjh_name &"</font>的手跳起了笨笨的探戈，体力和魅力下降1000。"
				mytl=-1000
				myml=-1000
			case 3,7
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& sjjh_name &"</font>鼓起勇气，面红红的走向<font color=blue>"& fromname &"</font>，托起她（他）的小手翩翩跳起了酒醉的探戈，两人终于在这样的意境中加深了友谊。"
			case 4
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>看到<font color=blue>"& sjjh_name &"</font>在侨都酒店舞厅里这样的不自然，就勇敢大方地主动邀他跳上一曲悠扬的探戈，两人陶醉在这夜的浪漫之中。道德不良，扣500道德"
				mydd=-500
			case 5
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>看到<font color=blue>"& sjjh_name &"</font>在侨都酒店舞厅里这样的不自然，就勇敢大方地主动邀他跳上一曲悠扬的探戈，两人陶醉在这夜的浪漫之中。体力长200，魅力长1000，银两多五万"
				mytl=200
				myml=1000
				yingliang=yingliang+50000
			case 6
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>看到<font color=blue>"& sjjh_name &"</font>在侨都酒店舞厅里这样的不自然，就勇敢大方地主动邀他跳上一曲悠扬的探戈，晕。跳舞除我之外，你别再找谁了，阿拉是高手，两人陶醉在这夜的浪漫之中。"
			case else
				tw="<bgsound src=mid/tg.mid loop=1><img src=img/tg.gif><font color=blue>"& fromname &"</font>看到<font color=blue>"& sjjh_name &"</font>在侨都酒店舞厅里这样的不自然，就勇敢大方地主动邀他跳上一曲悠扬的探戈，两人陶醉在这夜的浪漫之中。"
		end select
	case "迪斯高"
		randomize timer
		r=int(rnd*9)+1
		select case r
			case 1
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>套着一片草裙疯狂地扭动着腰肢，还不时地向大家抛几个飞吻，弄的聊天室里怪叫连连，<font color=blue>"& fromname &"</font>也疯狂地扭动着那肥大的臀部。"
			case 2,8
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>套着一片草裙疯狂地扭动着纤纤细肢，还不时地向大家抛几个媚眼，弄的聊天室里怪叫连连，<font color=blue>"& fromname &"</font>也疯狂地伴着她身边。"
			case 3,7
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& fromname &"</font>套着一片草裙疯狂弄的跳起DISCO，聊天室里怪叫连连，<font color=blue>"& sjjh_name &"</font>也疯狂地合着起来，SO IN。加五万银两"
				yingliang=yinliang+50000
			case 4
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>和<font color=blue>"& fromname &"</font>在迪厅疯狂领舞，跳反了，体力内力下降500."
				mytl=-500
				mynl=-500
				yingliang=0
			case 5
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>和<font color=blue>"& fromname &"</font>在迪厅疯狂领舞，音乐震耳，灯光闪耀，台下是挥舞着的手的海洋.加银两十万"
				yingliang=yingliang+100000
			case 6
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>和<font color=blue>"& fromname &"</font>在迪厅疯狂领舞，谢庭峰也不过如此，哈哈，谁敢与我俩一较高低."
			case else
				tw="<bgsound src=mid/disco.mid loop=1><img src=img/disco.gif><font color=blue>"& sjjh_name &"</font>和<font color=blue>"& fromname &"</font>在红绿灯酒巴里，音乐震耳，灯光闪耀，真是很得意呀，酒得“梅”人归，跳得红粉笑.加体力500"
				mytl=500
		end select
end select

conn.execute "update 用户 set 银两=银两-"&yinliang&",体力=体力+"&mytl&",内力=内力+"&mynl&",魅力=魅力+"&myml&",allvalue=allvalue+"&myallvalue&",道德=道德+"&mydd&" where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两+"&yinliang&",体力=体力-800,内力=内力-1000 where 姓名='"&fromname&"'"
set rs=nothing
conn.close
set conn=nothing
tw=ab&tw&wei
Application.Lock
says="<font color=red><b>【侨都舞厅】</b></font>"&tw	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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