<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<9 then
	Response.Write "<script language=JavaScript>{alert('提示：你的等级不够,你不能操作!');window.close();}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<5 then
	s=5-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你刚放过怪,请再等["&s&"]分钟再来！');location.href ='javascript:history.go(-1)';</script>"
	Response.End
end if
conn.Execute "update 用户 set 操作时间=now() where 姓名='" & aqjh_name &"'"
set rs=nothing
conn.close
set conn=nothing
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize
s=int(rnd*value)+1
yn=0
select case cz
case "狗狗"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/gogo.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
	fafang="<bgsound src=wav/dog.wav loop=2><font color=red>【消息】突然间一声狗叫……“狗狗呀！”江湖里面怕狗的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头狗狗闯进聊天室，狗狗体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "衰哥"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/boy.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
	fafang="<bgsound src=wav/Ye150.wav loop=3><font color=red>【消息】一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "武士"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/ws.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FX40A.wav loop=2><font color=red>【消息】哇！江湖里来了一个武士，要找大家比武，你们谁想试试？武士体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "八国联军"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/lj.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
	fafang="<bgsound src=wav/zha.wav loop=2><font color=red>【消息】八国联军入侵中国，江湖儿女们杀鬼子呀！冲呀~~~~~~~~~~~~八国联军人数："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "人贩"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/rf.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.wav loop=1><font color=#FF6600>【江湖任务】捉拿人贩子:叶三娘(ye sanniang)：“可恶的人贩子，把我那可怜的孩儿拐到哪里去啦？要是再找不到的话，我就见一个孩子杀一个。”人贩子体力：+"&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "小偷"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/piaoke.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.wav loop=1><font color=#FF6600>【江湖任务】抓小偷任务开始：这年头江湖也不平静啊！居然有位贼头贼脑的小偷闯入江湖，谁把他逮住必有重谢！”</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "舞女"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/mm.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>【消息】咦哈~~~~~~江湖来了位漂亮的领舞小姐“会摇的你就好好摇，不会摇的爱怎么摇就怎么摇！”，舞女体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "烟花"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/yianhua.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/yianhua.GIF' border='0'></a>"
	fafang="<bgsound src=wav/yianhua.wma loop=1><font color=red>【消息】咦哈~~~~~~江湖烟花汇演，快看呀，呵呵，看烟花得要跑上楼顶，要体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/yianhua.wma loop=1>"&abc&"</marquee>"
case "狮子"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/xu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/xu.GIF' border='0'></a>"
	fafang="<bgsound src=wav/shizi.wav loop=1><font color=red>【消息】咦哈~~~~~~江湖请了舞狮来贺年了，你也可以来舞狮呀，不过，要体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/xu.wma loop=1>"&abc&"</marquee>"
case "非典"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/fd.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/fd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/feidian.wav loop=1><font color=red>【消息】美帝国主义把非典带给了江湖，为了严防非典，请大家努力杀毒，感染力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "土匪"
yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='tuhu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/jiao.gif' border='0'></a>"
    fafang="<bgsound src=wav/tufei.wav loop=1><font color=red>【消息】突然一群土匪闯入江湖抢劫，请高手们快去剿匪啊。土匪体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "苹果"
	yn=0
	Application.Lock
	Application("aqjh_yb1")=s+100
	Application.UnLock
	abc="<a href='fafang/apple.asp?tl="&Application("aqjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
	fafang="<bgsound src=wav/baguo.wav loop=1><font color=red>【消息】</font><font color=red>突然间从天上掉下来一个大苹果，吃了增加："&s+100&"体力![操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "饿鼠"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/shu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
	fafang="<bgsound src=wav/eshu.wav loop=1><font color=red>【消息】</font><font color=red>突然间江湖创入一只饿疯的老鼠……“饿鼠呀！”江湖里面怕鼠的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头饿鼠闯进聊天室，饿鼠体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "飞鸟"
yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='DABIAN.ASP?tl="&Application("aqjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
    fafang="<bgsound src=wav/niao.wav loop=2><font color=red>【消息】一只怪鸟从冥王星飞临江湖聊天室，太空鸟啊，一定会有好东西的，大家快打啊!怪鸟体力"&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
case "舞男"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/wn.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>【消息】江湖为女孩子们带来了最帅的领舞先生，举起你们的双手“摇起来~~~~~~”，舞男体力："&s&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "流氓"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/liumang.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
	fafang="<bgsound src=wav/liumang.wav loop=1><font color=red>【消息】一个猪哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。流氓猪哥体力："&s+100&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "老虎"
	yn=0
	Application.Lock
	Application("aqjh_dalie")="老虎"
	Application.UnLock
	fafang="<bgsound src=wav/hu.wav loop=1><font color=red>【消息】</font><font color=red>突然一只野兽<img src=img/laohu.gif>窜入江湖中伤人，请高手们快去打猎啊。[操作:"&aqjh_name&"]</font>"
case "僵尸"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/kl.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>"
	fafang="<bgsound src=wav/gui.wav loop=3><font color=red>【消息】突然间阴风阵阵……“僵尸呀！一个女子发出一声恐怖的尖叫，一个僵尸闯进聊天室，僵尸体力："&s+100&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "手枪"
	yn=0
	Application.Lock
	Application("aqjh_qi")=s+100
	Application.UnLock
	abc="<a href='fafang/qi.asp?tl="&Application("aqjh_qi")&"' target='d'><img src='img/tank.GIF' border='0'></a>"
	fafang="<bgsound src=wav/dz.wav loop=8><font color=red>【消息】</font><font color=red>一把精制手枪闯进江湖，中原武林人士看了目瞪口呆，也不知道这是什么东西。都先打了在说，精制手枪体力："&s+100&"[操作:"&aqjh_name&"]</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
end select
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
if yn=1 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	for nowinroom=0 to chatroomnum
		useronlinename=Application("aqjh_useronlinename"&nowinroom)
		online=Split(Trim(useronlinename)," ",-1)
		x=UBound(online)
		for i=0 to x
			conn.Execute "update 用户 set "&cz&"="&cz&"+" & s & " where 姓名='" & online(i) & "'"
		next
	next
	conn.close
	set conn=nothing
end if
says=fafang
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
	for nowinroom=0 to chatroomnum
		saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
		addmsg saystr
	next
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
Response.Redirect "../../ok.asp?id=705"
%>