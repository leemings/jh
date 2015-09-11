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
if aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('提示：这一些只有十级管理才可以操作!');window.close();}</script>"
	Response.End 
end if
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize
s=int(rnd*value)+1
yn=0
select case cz
case "银两"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发钱】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>向聊天室的每个人发了<font color=blue>+"& s &"</font>银两！"
case "金币"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发金币】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>取出金币，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>个!"
case "道德"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【传授道德】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>今天特别高兴，向聊天室的每个人传授了<font color=blue>+"& s &"</font>点道德!</font>"
case "魅力"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【传授魅力】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>向聊天室的每个人传授了<font color=blue>+"& s &"</font>点魅力!</font>"
case "法力"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发法力】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点法力!</font>"
case "轻功"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发轻功】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点轻功!</font>"
case "金"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发金属性】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点金属性!></font>"
case "木"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发木属性】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点木属性!</font>"
case "水"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发水属性】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点水属性!</font>"
case "火"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发火属性】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点火属性!></font>"
case "土"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发土属性】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>为了鼓励大家，给聊天室中“美”(每)人发放<font color=blue>+"& s &"</font>点土属性!</font>"
case "抢金卡"
	yn=0
        Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/jk.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
	fafang="<bgsound src=wav/py.mid loop=1><font color=red>【抢金卡】大家看看这是什么呀，哇,原来是金卡“"&s&"块”!大家快抢啊,谁抢到是谁的!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "抢存点"
	yn=0
	Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/cd.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/cd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/lw.mid loop=1><font color=red>【抢存点】</font><font color=red>站长今日高兴，将存点“"&s&"点”撒向空中!谁抢到是谁的！</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "抢豆点"
	yn=0
	Application.Lock
	Application("aqjh_yb")=s
	Application.UnLock
	abc="<a href='fafang/dd.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/dd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/MYFAVOR.mid loop=1><font color=red>【抢豆点】</font><font color=red>站长今日高兴，将豆点“"&s&"点”撒向空中!谁抢到是谁的！</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "泡豆点数"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【发豆子】</font><font color=blue>"& aqjh_name &"</font><font color=#ff0000>从生产队大院里出来，扛了一大麻袋的豆子，要去赶集去。谁知袋子漏了一路上豆子掉了<font color=blue>+"& s &"</font>个，嘿嘿。。见者有份！</font>"
case "狗狗"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/gogo.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
	fafang="<bgsound src=wav/dog.wav loop=2><font color=red>【消息】突然间一声狗叫……“狗狗呀！”江湖里面怕狗的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头狗狗闯进聊天室，狗狗体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "衰哥"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/boy.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
	fafang="<bgsound src=wav/Ye150.wav loop=3><font color=red>【消息】一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "武士"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/ws.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
	fafang="<bgsound src=wav/FX40A.wav loop=2><font color=red>【消息】哇！江湖来了一个武士，要找大家比武，你们谁想试试？武士体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
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
case "八国联军"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/lj.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
	fafang="<bgsound src=wav/zha.wav loop=2><font color=red>【消息】八国联军入侵江湖，江湖儿女们杀鬼子呀！冲呀~~~~~~~~~~~~八国联军人数："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "舞女"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/mm.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>【消息】咦哈~~~~~~江湖来了位漂亮的领舞小姐“会摇的你就好好摇，不会摇的爱怎么摇就怎么摇！”，舞女体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "烟花"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/yianhua.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/yianhua.GIF' border='0'></a>"
	fafang="<bgsound src=wav/yianhua.wma loop=1><font color=red>【消息】咦哈~~~~~~江湖烟花汇演，快看呀，呵呵，看烟花得要跑上楼顶，要体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/yianhua.wma loop=1>"&abc&"</marquee>"
case "狮子"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/xu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/xu.GIF' border='0'></a>"
	fafang="<bgsound src=wav/shizi.wav loop=1><font color=red>【消息】咦哈~~~~~~江湖请了舞狮来贺年了，你也可以来舞狮呀，不过，要体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/xu.wma loop=1>"&abc&"</marquee>"
case "非典"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/fd.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/fd.GIF' border='0'></a>"
	fafang="<bgsound src=wav/feidian.wav loop=1><font color=red>【消息】美帝国主义把非典带给了江湖，为了严防非典，请大家努力杀毒，感染力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "土匪"
yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/tuhu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/jiao.gif' border='0'></a>"
    fafang="<bgsound src=wav/tufei.wav loop=1><font color=red>【消息】突然一群土匪闯入江湖抢劫，请高手们快去剿匪啊。土匪体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "苹果"
	yn=0
	Application.Lock
	Application("aqjh_yb1")=s+100
	Application.UnLock
	abc="<a href='fafang/apple.asp?tl="&Application("aqjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
	fafang="<bgsound src=wav/baguo.wav loop=1><font color=red>【消息】</font><font color=red>突然间从天上掉下来一个大苹果，吃了增加："&s+100&"体力!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "饿鼠"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/shu.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
	fafang="<bgsound src=wav/eshu.wav loop=2><font color=red>【消息】</font><font color=red>突然间江湖创入一只饿疯的老鼠……“饿鼠呀！”江湖里面怕鼠的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头饿鼠闯进聊天室，饿鼠体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "飞鸟"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/DABIAN.ASP?tl="&Application("aqjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
    fafang="<bgsound src=wav/niao.wav loop=2><font color=red>【消息】一只怪鸟从冥王星飞临江湖聊天室，太空鸟啊，一定会有好东西的，大家快打啊!怪鸟体力"&s&"</font><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
case "舞男"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s
	Application.UnLock
	abc="<a href='fafang/wn.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
	fafang="<bgsound src=wav/tw.wav loop=2><font color=red>【消息】江湖为女孩子们带来了最帅的领舞先生，举起你们的双手“摇起来~~~~~~”，舞男体力："&s&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "流氓"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/liumang.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
	fafang="<bgsound src=wav/liumang.wav loop=1><font color=red>【消息】一个猪哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。流氓猪哥体力："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "卡片"
	yn=0
	Application.Lock
	Application("aqjh_kp")=s+1
	Application.UnLock
	abc="<a href='fafang/kp.asp?tl="&Application("aqjh_kp")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
	fafang="<font color=red>【消息】钟声是我的问候，歌声是我的祝福，雪花是我的贺卡，美酒是我的飞吻，清风是我的拥抱，快乐是我的礼物！统统都送给你，一张金光闪闪的卡片掉在江湖路上，各位大小虾快抢啊："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/kp.mid loop=1>"&abc&"</marquee>"
case "补药"
	yn=0
	Application.Lock
	Application("aqjh_by")=s+1
	Application.UnLock
	abc="<bgsound src=wav/phant030a.wav loop=1><a href='fafang/by.asp?tl="&Application("aqjh_by")&"' target='d'><img src='img/by.GIF' border='0'></a>"
	fafang="<font color=red>【消息】不知道谁的药掉在路上了："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "配药"
	yn=0
	Application.Lock
	Application("aqjh_py")=s+1
	Application.UnLock
	abc="<a href='fafang/py.asp?tl="&Application("aqjh_py")&"' target='d'><img src='img/py.GIF' border='0'></a>"
	fafang="<font color=red>【消息】财神除夕到，奖金满钱包；开春撞桃花，美女入怀抱；盛夏日炎炎，欧美任逍遥；金秋重阳节，仕途步步高！有好东东检了，谁捡到是谁的："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=45><bgsound src=wav/phant030a.wav loop=1>"&abc&"</marquee>"
case "礼物"
	yn=0
	Application.Lock
	Application("aqjh_lw")=s+1
	Application.UnLock
	abc="<a href='fafang/lw.asp?tl="&Application("aqjh_lw")&"' target='d'><img src='img/lw.GIF' border='0'></a>"
	fafang="<font color=red>【消息】喜庆佳节，祝君出门遇贵人，在家听喜报！收红包了："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/LW.mid loop=1>"&abc&"</marquee>"
case "月饼"
	yn=0
	Application.Lock
	Application("aqjh_zqyb")=s+1
	Application.UnLock
	abc="<a href='fafang/zqyb.asp?tl="&Application("aqjh_zqyb")&"' target='d'><img src='img/zqyb.GIF' border='0'></a>"
	fafang="<font color=red>【消息】迎中秋，送大礼，我们的站长给大家送月饼来了，吃了可以补："&s+100&"体力</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/LW.mid loop=1>"&abc&"</marquee>"
case "老虎"
	yn=0
	Application.Lock
	Application("aqjh_dalie")="老虎"
	Application.UnLock
	fafang="<bgsound src=wav/hu.wav loop=1><font color=red>【消息】</font><font color=red>突然一只野兽<img src=img/laohu.gif>窜入江湖中伤人，请高手们快去打猎啊。</font>"
case "僵尸"
	yn=0
	Application.Lock
	Application("aqjh_kl")=s+100
	Application.UnLock
	abc="<a href='fafang/kl.asp?tl="&Application("aqjh_kl")&"' target='d'><img src='img/gui.GIF' border='0'></a>"
	fafang="<bgsound src=wav/gui.wav loop=3><font color=red>【阴讯】突然间阴风阵阵……“僵尸呀！一个女子发出一声恐怖的尖叫，一个僵尸闯进聊天室，僵尸体力："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "元宝"
	yn=0
	Application.Lock
	Application("aqjh_yb")=s+100
	Application.UnLock
	abc="<a href='fafang/yb.asp?tl="&Application("aqjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
	fafang="<bgsound src=wav/zt.mid loop=1><font color=red>【消息】天上掉下一个金元宝,谁抢到是谁的!价值："&s+100&"两!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "蛋糕"
	yn=0
	Application.Lock
	Application("aqjh_yb1")=s+100
	Application.UnLock
	abc="<a href='fafang/yb1.asp?tl="&Application("aqjh_yb1")&"' target='d'><img src='img/1-2.gif' border='0'></a>"
	fafang="<bgsound src=wav/11.mid loop=1><font color=red>【消息】江湖生日到了。站长"& aqjh_name &"谢谢各位了，大家吃蛋糕，吃了增加："&s+100&"体力!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "别墅"
	yn=0
	Application.Lock
	Application("aqjh_yb1")=s+1000
	Application.UnLock
	abc="<a href='fafang/bieshu.asp?tl="&Application("aqjh_yb1")&"' target='d'><img src='img/bieshu.gif' border='0'></a>"
	fafang="<font color=red>【消息】</font><font color=red>江湖为了感谢朋友们的支持和厚爱，送豪华别墅一栋，得到加："&s+1000&"体力!</font><br><marquee width=100% behavior=alternate scrollamount=15><bgsound src=wav/xu.wma loop=1>"&abc&"</marquee>"
case "站长"
	yn=0
	Application.Lock
	Application("aqjh_hj")=s+100
	Application.UnLock
	abc="<a href='fafang/hj.asp?tl="&Application("aqjh_hj")&"' target='d'><img src='img/hj.gif' border='0'></a>"
	fafang="<bgsound src=wav/FAQIAN.wav loop=1><font color=red>【消息】讨人喜欢的站长"& aqjh_name &"进江湖啦，大家快亲啊：亲了奖励"&s+100&"体力!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "体力"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【体力】</font>大侠<font color=blue>"& aqjh_name &"</font><font color=#ff0000>最近很高兴，路过福缘客栈，请所有大侠吃饭,大侠们体力增加：<font color=blue>+"& s &"</font>点!！</font>"
case "内力"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【内力】</font>大侠<font color=blue>"& aqjh_name &"</font><font color=#ff0000>路过江湖演武场，不忍看到江湖兄弟们辛苦修练，便传授一招任我行的吸星大法，谁知却被各位大虾每人吸走内力:<font color=blue>+"& s &"</font>，唉！好心没有好报~~</font>"
case "武功"
	yn=1
	fafang="<bgsound src=wav/zt.mid loop=1><font color=green>【武功】</font>大侠<font color=blue>"& aqjh_name &"</font><font color=#ff0000>路经此地，给聊天室中每人传授武功：<font color=blue>+"& s &"</font>点~~~</font>"
case "手枪"
	yn=0
	Application.Lock
	Application("aqjh_qi")=s+100
	Application.UnLock
	abc="<a href='fafang/qi.asp?tl="&Application("aqjh_qi")&"' target='d'><img src='img/tank.GIF' border='0'></a>"
	fafang="<bgsound src=wav/dz.wav loop=8><font color=red>【消息】</font><font color=red>一把精制手枪闯进江湖，中原武林人士看了目瞪口呆，也不知道这是什么东西。都先打了在说，精制手枪体力："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=25>"&abc&"</marquee>"
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