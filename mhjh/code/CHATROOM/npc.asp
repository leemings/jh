<%
case 1
Application("yx8_mhjh_klt1")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>突然一只神鹿来偷袭<font color=#0000FF>" & username & "</font>,吸取<font color=#0000FF>" & username & "</font>体力"&Application("yx8_mhjh_klt1")*5000&"</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc1.asp?r="&Application("yx8_mhjh_klt1")&" target=optfrm><img src=img/tr003.gif  border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt1")*5000&" where 姓名='" & username & "'")

case 2
	mess="<font color=green>【仙人指路】</font><font color=red>江湖中又一位大侠修成正果，巧遇仙人，南方出现一道绚丽的彩虹，彩虹下金光灿烂，里面藏着奇珍异宝，谁能拿到，定然受益菲浅！但前途坎坷，没有"&jstl&"的体力是无法成功的！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=mohuanshen.asp?r="&Application("yx8_mhjh_klt&r")&" target=optfrm><img src=img/mon18.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
case 3
Application("yx8_mhjh_klt3")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>突然一只庞大的恐怖飞兽闯进江湖，咬伤" & username & "体力"&Application("yx8_mhjh_klt3")*2000&"！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc2.asp?r="&Application("yx8_mhjh_klt3")&" target=optfrm><img src=img/mon24.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt3")*2000&" where 姓名='" & username & "'")
case 4
	mess="<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center>皇城公告</td> </tr> <tr><td valign=center align=center >快乐江湖的动武时间是每小时的后30分钟；前30分钟可以自由打怪物，同样能增加您的状态；但一定要小心啊，别被怪物咬死，即使你开启了保护也无法抵挡怪物的袭击！所以，一定要给自己加一些体力和内力，内力没了也一样要死人的！现在泡分得状态很少,主要依靠不断的修炼来提高自己的各种状态.魔幻之神不欢迎懒汉哦,呵呵</td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table>"
case 5
Application("yx8_mhjh_klt4")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>聊天室突然一阵震荡，一只巨大的尸王蹒跚而来，吸取<font color=#0000FF>" & username & "</font>内力"&Application("yx8_mhjh_klt4")*1000&"！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("yx8_mhjh_klt4")&" target=optfrm><img src=img/mon22.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力-"&Application("yx8_mhjh_klt4")*1000&" where 姓名='" & username & "'")
case 6
Application("yx8_mhjh_klt4")=r
	mess="<font color=green>【怪物呈祥】</font><font color=red>一阵祥和的风儿吹来，一只修炼千年的神灯翩然而至，给<font color=#0000FF>" & username & "</font>增加积分"&Application("yx8_mhjh_klt4")*20&"！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc3.asp?r="&Application("yx8_mhjh_klt4")&" target=optfrm><img src=img/mon20.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 积分=积分+"&Application("yx8_mhjh_klt4")*20&" where 姓名='" & username & "'")
case 7
	mess="<font color=green>【怪物偷袭】</font><font color=red>一股凄厉的寒风横扫江湖，寒冰卡比横行肆虐，大侠<font color=#0000FF>" & username & "</font>被咬去"&Application("yx8_mhjh_klt6")*100&"点防御！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("yx8_mhjh_klt6")&" target=optfrm><img src=img/mon15.gif   border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 防御=防御-"&Application("yx8_mhjh_klt6")*100&" where 姓名='" & username & "'")
case 8
Application("yx8_mhjh_klt6")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>一股热浪突袭江湖，火龙闯入一群若女子当中，见人就咬，<font color=#0000FF>" & username & "</font>被咬伤体力"&Application("yx8_mhjh_klt6")*2000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc5.asp?r="&Application("yx8_mhjh_klt6")&" target=optfrm><img src=img/mon23.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt6")*2000&" where 姓名='" & username & "'")
case 9
Application("yx8_mhjh_klt7")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>金奥克在深山被一股野火驱赶下山，游荡进快乐江湖，饥不择食，吸收<font color=#0000FF>" & username & "</font>道德"&Application("yx8_mhjh_klt7")*500&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc7.asp?r="&Application("yx8_mhjh_klt7")&" target=optfrm><img src=img/mon28.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 道德=道德-"&Application("yx8_mhjh_klt7")*500&" where 姓名='" & username & "'")
case 10
Application("yx8_mhjh_klt8")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>魔幻剑魂在哥特山顶苦心修炼，一股剑气伤了<font color=#0000FF>" & username & "</font>资质"&Application("yx8_mhjh_klt8")*100&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc6.asp?r="&Application("yx8_mhjh_klt8")&" target=optfrm><img src=img/mon27.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 资质=资质-"&Application("yx8_mhjh_klt8")*100&" where 姓名='" & username & "'")
case 11
Application("yx8_mhjh_klt9")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>哇~~~好酷的一只鸭!真是林子大了，什么鸟都有，一只蹩脚鸭跑进快乐江湖，偷了<font color=#0000FF>" & username & "</font>白银"&Application("yx8_mhjh_klt9")*5000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc8.asp?r="&Application("yx8_mhjh_klt9")&" target=optfrm><img src=img/mon3.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两-"&Application("yx8_mhjh_klt9")*5000&" where 姓名='" & username & "'")
case 12
Application("yx8_mhjh_klt10")=r
	mess="<font color=green>【怪物偷袭】</font><font color=red>闪电龙突然冲进快乐江湖，一道霹雳把<font color=#0000FF>" & username & "</font>打得浑身黑糊糊的，并夺取积分"&Application("yx8_mhjh_klt10")*10&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc9.asp?r="&Application("yx8_mhjh_klt10")&" target=optfrm><img src=img/mon4.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 积分=积分-"&Application("yx8_mhjh_klt10")*10&" where 姓名='" & username & "'")
case 13
Application("yx8_mhjh_klt11")=r
	mess="<font color=green>【怪物呈祥】</font><font color=red>一只百年不遇的江湖之宝梦精灵突然出现在聊天室内，送给<font color=#0000FF>" & username & "</font>体力"&Application("yx8_mhjh_klt11")*4000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc10.asp?r="&Application("yx8_mhjh_klt11")&" target=optfrm><img src=img/mon10.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力+"&Application("yx8_mhjh_klt11")*4000&" where 姓名='" & username & "'")
case 14
Application("yx8_mhjh_klt12")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>一只火史莱姆蹦蹦跳跳闯进快乐江湖，向<font color=#0000FF>" & username & "</font>挑衅似的吸取内力"&Application("yx8_mhjh_klt12")*1000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc11.asp?r="&Application("yx8_mhjh_klt12")&" target=optfrm><img src=img/mon9.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力-"&Application("yx8_mhjh_klt12")*1000&" where 姓名='" & username & "'")
case 15
Application("yx8_mhjh_klt13")=r
	mess="<font color=green>【鬼魂附体】</font><font color=red>大侠<font color=#0000FF>" & username & "</font>修炼内功时不小心走火入魔，被刺球宝宝的魂灵附着到身上，被咬伤"&Application("yx8_mhjh_klt13")*4000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc12.asp?r="&Application("yx8_mhjh_klt13")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt13")*4000&" where 姓名='" & username & "'")
case 16
Application("yx8_mhjh_klt14")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只怪物闯进聊天室!杀伤" & username & "体力"&Application("yx8_mhjh_klt14")*3000&"点！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc4.asp?r="&Application("yx8_mhjh_klt14")&" target=optfrm><img src=img/mon8.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt14")*3000&" where 姓名='" & username & "'")
case 17
Application("yx8_mhjh_klt15")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>好漂亮啊~~~~一个不知道什么东西变的舞男翩跹起舞,兴奋中丢失了一个风铃,被" & username & "得到！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc15.asp?r="&Application("yx8_mhjh_klt15")&" target=optfrm><img src=img/mon12.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 风铃=风铃+1 where 姓名='" & username & "'")
case 18
Application("yx8_mhjh_klt16")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>好漂亮啊！一个不知道什么东西变的舞男翩跹起舞,兴奋中丢失了一个风铃,被" & username & "得到！</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc16.asp?r="&Application("yx8_mhjh_klt16")&" target=optfrm><img src=img/mon12.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 风铃=风铃+1 where 姓名='" & username & "'")
case 19
Application("yx8_mhjh_klt17")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,吸取" & username & "内力"&Application("yx8_mhjh_klt17")*500&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc17.asp?r="&Application("yx8_mhjh_klt17")&" target=optfrm><img src=img/Shenmo.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 内力=内力-"&Application("yx8_mhjh_klt17")*500&" where 姓名='" & username & "'")
case 20
Application("yx8_mhjh_klt18")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,打掉" & username & "攻击"&Application("yx8_mhjh_klt18")*10&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc18.asp?r="&Application("yx8_mhjh_klt18")&" target=optfrm><img src=img/Ying.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 攻击=攻击-"&Application("yx8_mhjh_klt18")*10&" where 姓名='" & username & "'")
case 21
Application("yx8_mhjh_klt19")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,打掉" & username & "防御"&Application("yx8_mhjh_klt19")*10&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc19.asp?r="&Application("yx8_mhjh_klt&r")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 防御=防御-"&Application("yx8_mhjh_klt19")*10&" where 姓名='" & username & "'")
case 21
Application("yx8_mhjh_klt20")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,抢走" & username & "银两"&Application("yx8_mhjh_klt20")*1200&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc20.asp?r="&Application("yx8_mhjh_klt20")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两-"&Application("yx8_mhjh_klt20")*1200&" where 姓名='" & username & "'")
case 22
Application("yx8_mhjh_klt21")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,偷走" & username & "积分"&Application("yx8_mhjh_klt21")*5&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc21.asp?r="&Application("yx8_mhjh_klt21")&" target=optfrm><img src=img/Ying1.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 积分=积分-"&Application("yx8_mhjh_klt21")*5&" where 姓名='" & username & "'")
case 23
Application("yx8_mhjh_klt22")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,降低" & username & "道德"&Application("yx8_mhjh_klt22")*500&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc22.asp?r="&Application("yx8_mhjh_klt22")&" target=optfrm><img src=img/tr003.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 道德=道德-"&Application("yx8_mhjh_klt22")*500&" where 姓名='" & username & "'")
case 24
Application("yx8_mhjh_klt23")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,杀伤" & username & "体力"&Application("yx8_mhjh_klt23")*1900&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc23.asp?r="&Application("yx8_mhjh_klt23")&" target=optfrm><img src=img/tr002.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt23")*1900&" where 姓名='" & username & "'")
case 25
Application("yx8_mhjh_klt24")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,杀伤" & username & "体力"&Application("yx8_mhjh_klt24")*2300&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc14.asp?r="&Application("yx8_mhjh_klt24")&" target=optfrm><img src=img/tr004.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt24")*2300&" where 姓名='" & username & "'")
case 26
Application("yx8_mhjh_klt25")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,杀伤" & username & "体力"&Application("yx8_mhjh_klt25")*2000&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc25.asp?r="&Application("yx8_mhjh_klt25")&" target=optfrm><img src=img/tr103.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt25")*2000&" where 姓名='" & username & "'")
case 27
Application("yx8_mhjh_klt26")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,杀伤" & username & "体力"&Application("yx8_mhjh_klt26")*1400&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc26.asp?r="&Application("yx8_mhjh_klt26")&" target=optfrm><img src=img/tr102.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt26")*1400&" where 姓名='" & username & "'")
case 28
Application("yx8_mhjh_klt27")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,杀伤" & username & "体力"&Application("yx8_mhjh_klt27")*1100&"点</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc27.asp?r="&Application("yx8_mhjh_klt27")&" target=optfrm><img src=img/tr101.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 体力=体力-"&Application("yx8_mhjh_klt27")*1100&" where 姓名='" & username & "'")
case 29
Application("yx8_mhjh_klt28")=r
	mess="<font color=green>【怪物袭击】</font><font color=red>突然一只<font color=#0000FF>怪物</font>闯进来,偷走" & username & "银两"&Application("yx8_mhjh_klt28")*3000&"两</font><br><marquee width=100% behavior=alternate scrollamount=15><a href=npc/npc28.asp?r="&Application("yx8_mhjh_klt28")&" target=optfrm><img src=img/tr002.gif border=0></a></marquee><bgsound src=../mid/Ye150.wav loop=3>"
Set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst = server.createobject("adodb.recordset") 
    conn.execute("update 用户 set 银两=银两-"&Application("yx8_mhjh_klt28")*3000&" where 姓名='" & username & "'")
case 30
	mess="<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center>皇城公告</td> </tr> <tr><td valign=center align=center >快乐江湖的动武时间是每小时的后30分钟；前30分钟可以自由打怪物，同样能增加您的状态；但一定要小心啊，别被怪物咬死，即使你开启了保护也无法抵挡怪物的袭击！所以，一定要给自己加一些体力和内力，内力没了也一样要死人的,现在泡分得状态很少,主要依靠不断的修炼来提高自己的各种状态.魔幻之神不欢迎懒汉哦,呵呵！</td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table>"
end select
conn.close
set conn=nothing
application.unlock
dati=Application("yx8_mhjh_talkarr")
		talkpoint=clng(Application("yx8_mhjh_talkpoint"))
		dim newdati(600)
		j=1
		for i=11 to 600 step 10
			newdati(j)=dati(i)
			newdati(j+1)=dati(i+1)
			newdati(j+2)=dati(i+2)
			newdati(j+3)=dati(i+3)
			newdati(j+4)=dati(i+4)
			newdati(j+5)=dati(i+5)
			newdati(j+6)=dati(i+6)
			newdati(j+7)=dati(i+7)
			newdati(j+8)=dati(i+8)
			newdati(j+9)=dati(i+9)
			j=j+10
		next
		newdati(591)=talkpoint+1
		newdati(592)=1
		newdati(593)=0
		newdati(594)=""
		newdati(595)="大家"
		newdati(596)=""
		newdati(597)=namecolor
		newdati(598)=wordcolor
		newdati(599)=""&mess&""
		newdati(600)=chatroomsn
		Application.Lock
		Application("yx8_mhjh_talkpoint")=talkpoint+1
		Application("yx8_mhjh_talkarr")=newdati
		Application.UnLock
		erase newdati
		erase dati
end if
end if
%>