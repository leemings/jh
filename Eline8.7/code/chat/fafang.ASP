<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 then
	Response.Write "<script language=JavaScript>{alert('提示：这一些只有站长才可以操作!');window.close();}</script>"
	Response.End 
end if
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
s=int(rnd*value)+1
yn=0
select case cz
case "银两"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发钱】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>发钱喽！发钱喽！<img src=img/251.gif><img src=img/251.gif><img src=img/251.gif> <font color=blue>"& sjjh_name &"</font>给聊天室的每个人发了<font color=blue>+"& s &"</font>两金光闪闪的银子！</font></b><br><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif>"
case "金币"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【金币】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，从自家金库中取出大量金币，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>个金币!</font></b>"
case "银币"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【银币】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，从自家金库中取出大量银币，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>个银币!</font></b>"
case "卡片"
	yn=0
	Application.Lock
	Application("sjjh_kp")=s+1
	Application.UnLock
	abc="<a href='kp2.asp?tl="&Application("sjjh_kp")&"' target='d'><img src='img/kapian.GIF' border='0'></a>"
	fafang="<font color=red>一张金光闪闪的卡片掉在江湖路上，各位大小虾快抢啊!</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/kp.MID loop=1>"&abc&"</marquee>"
case "道德"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【道德】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，向聊天室的每个人传授了<font color=blue>+"& s &"</font>点道德!</font></b>"
case "魅力"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【魅力】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>请来美容师，给聊天室中每人美容，每人得到魅力<font color=blue>+"& s &"</font>点魅力!</font></b>"
case "法力"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【法力】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点法力!</font></b>"
case "轻功"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【轻功】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点轻功!</font></b>"
case "金"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【金属性】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点金属性!></font></b>"
case "木"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【木属性】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点木属性!</font></b>"
case "水"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【水属性】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点水属性!</font></b>"
case "火"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【火属性】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点火属性!></font></b>"
case "土"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【土属性】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>今天特别高兴，站长为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点土属性!</font></b>"
case "allvalue"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【存点】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>泡点是很难的，站长为了鼓励各位玩家泡点，每个人发放存点<font color=blue>+"& s &"</font>点，大家还要继续努力~~</font></b>"
case "泡豆点数"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【豆子】</font><font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>从生产队大院里出来，扛了一大麻袋的豆子，要去赶集去。谁知袋子漏了一路上豆子掉了<font color=blue>+"& s &"</font>个，嘿嘿。。见者有份！</font></b>"
case "老虎"
	yn=0
	Application.Lock
	Application("sjjh_dalie")="老虎"
	Application.UnLock
	fafang="<bgsound src=wav/hu.wav loop=3><font color=red>【消息】</font><b><font color=red>突然一只野兽<img src=img/laohu.gif>窜入江湖中伤人，请高手们快去打猎啊。</font></b>"
case "僵尸"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='kl.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/kl.GIF' border='0'></a>"
	fafang="<bgsound src=wav/gui.wav loop=3><font color=red>【消息】</font><b><font color=red>突然间阴风陈陈……“僵尸呀！”一女子尖叫，<bgsound src=wav/gui.wav loop=1>一头僵尸闯进聊天室，僵尸体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "狗狗"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='dog.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
	fafang="<bgsound src=wav/dog.WAV loop=3><font color=red>【消息】</font><b><font color=red>突然间一声狗叫……“狗狗呀！”江湖里面怕狗的人尖叫，<bgsound src=wav/gougou.wav loop=1>大叫到不要咬我，不要咬我，我怕怕！！！一头狗狗闯进聊天室，狗狗体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "饿鼠"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='shu.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
	fafang="<bgsound src=wav/shu.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间江湖创入一只饿疯的老鼠……“饿鼠呀！”江湖里面怕鼠的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头饿鼠闯进聊天室，饿鼠体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "怪鸟"
yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='DABIAN.ASP?tl="&Application("sjjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
    fafang="<bgsound src=wav/niao.wav loop=1><font color=red>【消息】</font><b><font color=red>一只怪鸟从冥王星飞临江湖聊天室，太空鸟啊，一定会有好东西的，大家快打啊!怪鸟体力"&s&"</font></b><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
case "衰哥"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='boy.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.WAV loop=3><font color=red>【消息】</font><b><font color=red>一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "狗狗"
	yn=0
	Application.Lock
	Application("ruhui_kl")=s
	Application.UnLock
	abc="<a href='dog.asp?tl="&Application("ruhui_kl")&"' target='d'><img src='img/dog.GIF' border='0'></a>"
	fafang="<bgsound src=wav/dog.WAV loop=3><font color=red>【消息】</font><b><font color=red>不知从哪冒出来一只小狗，<bgsound src=wav/gougou.wav loop=1>想吃狗肉的快打......”，狗狗体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "武士"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='ws.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
	fafang="<bgsound src=wav/bsj.wav loop=3><font color=red>【消息】</font><b><font color=red>哇！『快乐江湖』来了一个武士，要找大家比武，你们谁想试试？武士体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "八国联军"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='lj.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
	fafang="<bgsound src=wav/baguo.wav loop=3><font color=red>【消息】</font><b><font color=red>八国联军入侵『快乐江湖』，江湖儿女们杀鬼子呀！冲呀~~~~~~~~~~~~八国联军人数："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "舞女"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='mm.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
	fafang="<bgsound src=wav/wn2.wav loop=1><font color=red>【消息】</font><b><font color=red>咦哈~~~『快乐江湖』来了位漂亮的领舞小姐“会摇的你就好好摇，不会摇的爱怎么摇就怎么摇！”，舞女体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "土匪"
yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='tuhu.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/jiao.gif' border='0'></a>"
fafang="<bgsound src=wav/fei.WAV loop=3><font color=red>【消息】</font><b><font color=red>突然一群土匪闯入江湖抢劫，请高手们快去剿匪啊。土匪体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "舞男"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s
	Application.UnLock
	abc="<a href='wn.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
	fafang="<bgsound src=wav/wn1.wav loop=1><font color=red>【消息】</font><b><font color=red>站长为女孩子们带来了『快乐江湖』最帅的领舞先生，举起你们的双手“摇起来~~~”，舞男体力："&s&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "流氓"
	yn=0
	Application.Lock
	Application("sjjh_kl")=s+100
	Application.UnLock
	abc="<a href='liumang.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
	fafang="<bgsound src=wav/zhu.wav loop=1><font color=red>【消息】</font><b><font color=red>一个猪哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。流氓猪哥体力："&s+100&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "补药"
	yn=0
	Application.Lock
	Application("sjjh_by")=s+1
	Application.UnLock
	abc="<a href='by.asp?tl="&Application("sjjh_by")&"' target='d'><img src='img/by.GIF' border='0'></a>"
	fafang="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>不知道谁的药掉在路上了："&s+100&"</font></b><br><marquee width=100% behavior=alternate scrollamount=50>"&abc&"</marquee>"
case "配药"
	yn=0
	Application.Lock
	Application("sjjh_py")=s+1
	Application.UnLock
	abc="<a href='py2.asp?tl="&Application("sjjh_py")&"' target='d'><img src='img/py.GIF' border='0'></a>"
	fafang="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>炼丹没炼成，配药散落一地，将配药丢在路上，谁捡到是谁的!</font></b><br><marquee width=100% behavior=alternate scrollamount=45><bgsound src=wav/py.MID loop=1>"&abc&"</marquee>"
case "苹果"
	yn=0
	Application.Lock
	Application("sjjh_yb1")=s+100
	Application.UnLock
	abc="<a href='apple.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
	fafang="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间从天上掉下来一个大苹果，吃了增加："&s+100&"体力!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "别墅"
	yn=0
	Application.Lock
	Application("sjjh_yb1")=s+1000
	Application.UnLock
	abc="<a href='bieshu.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/bieshu.gif' border='0'></a>"
	fafang="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>『快乐江湖』为了感谢朋友们的支持和厚爱，送豪华别墅一栋，得到加："&s+1000&"体力!</font></b><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/HAPPY.MID loop=1>"&abc&"</marquee>"
case "礼物"
	yn=0
	Application.Lock
	Application("sjjh_lw")=s+1
	Application.UnLock
	abc="<a href='lw.asp?tl="&Application("sjjh_lw")&"' target='d'><img src='img/lw.jpg' border='0'></a>"
	fafang="<bgsound src=wav/zhanfa.wav loop=1><font color=red>【消息】</font><b><font color=red>站长为五湖四海的朋友献礼了，祝愿大家天天开心、万事如意！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "元宝"
	yn=0
	Application.Lock
	Application("sjjh_yb")=s
	Application.UnLock
	abc="<a href='yb.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
	fafang="<bgsound src=wav/diaoxia.WAV loop=1><font color=red>【消息】</font><b><font color=red>今天是什么日子，突然间从天上掉下来一个大元宝，价值："&s&"两!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "抢存点"
	yn=0
	Application.Lock
	Application("sjjh_yb")=s
	Application.UnLock
	abc="<a href='cd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/cd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/MYFAVOR.MID loop=1><font color=red>【消息】</font><b><font color=red>回首当年真伟大，把自己的存点“"&s&"点”撒向空中!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "抢豆点"
	yn=0
	Application.Lock
	Application("sjjh_yb")=s
	Application.UnLock
	abc="<a href='dd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/dd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/MYFAVOR.MID loop=1><font color=red>【消息】</font><b><font color=red>回首当年真伟大，把自己的豆点“"&s&"点”撒向空中!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "抢金卡"
	yn=0
	Application.Lock
	Application("sjjh_yb")=s
	Application.UnLock
	abc="<a href='jk.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
	fafang="<bgsound src=wav/MYFAVOR.MID loop=1><font color=red>【消息】</font><b><font color=red>回首当年真伟大，把自己的金卡“"&s&"块”撒向空中!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "体力"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【体力】</font>大侠<font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>最近发福了，路过福缘客栈，请所有大侠吃饭,大侠们体力增加：<font color=blue>+"& s &"</font>点!！</font></b>"
case "内力"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【内力】</font>大侠<font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>路过江湖演武场，不忍看到江湖兄弟们辛苦修练，便传授一招吸星大法，每人增加内力:<font color=blue>+"& s &"</font></b>"
case "武功"
	yn=1
	fafang="<bgsound src=wav/fa.wav loop=1><font color=green>【武功】</font>大侠<font color=blue>"& sjjh_name &"</font><B><font color=#ff0000>路经此地，给聊天室中每人发放武功：<font color=blue>+"& s &"</font>点~~~</font></b>"
case "手枪"
	yn=0
	Application.Lock
	Application("sjjh_qi")=s+100
	Application.UnLock
	abc="<a href='qi.asp?tl="&Application("sjjh_qi")&"' target='d'><img src='img/tank.GIF' border='0'></a>"
	fafang="<bgsound src=wav/dz.wav loop=8><font color=red>【消息】</font><b><font color=red>一把精制手枪闯进『快乐江湖』，中原武林人士看了目瞪口呆，也不知道这是什么东西。都先打了在说，精制手枪体力："&s+100&"</font></b><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"	
end select
if yn=1 then
	useronlinename=Application("sjjh_useronlinename"&nowinroom)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	for i=0 to x
		conn.Execute "update 用户 set "&cz&"="&cz&"+" & s & " where 姓名='" & online(i) & "'"
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
Response.Redirect "../ok.asp?id=705"
%>
