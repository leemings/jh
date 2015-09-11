<html>

<head>
<meta http-equiv="refresh" content="120;url=npc.asp">
<style type="text/css">
A {COLOR: #ffffff; TEXT-DECORATION: none}
td{font-size:9pt;color:ffffff}
</style>
</head>

<body background=bg.gif oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<!--#include file="npc_chat.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../showchat.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
ysjh_name=Session("ysjh_name")
ysjh_grade=Session("ysjh_grade")
ysjh_jhdj=Session("ysjh_jhdj")
nowinroom=session("nowinroom")
ysjh_roominfo=split(Application("ysjh_room"),";")
chatinfo=split(ysjh_roominfo(nowinroom),"|")
if ysjh_name="" then Response.Redirect "../error.asp?id=440"
allnpc=Application("ysjh_npc")
response.write now()
'判断刷新时间是否被卡
sxsj=DateDiff("n",application("npc_sxsj"),now())
if sxsj>10 then
	application("npc_sxsj")=now()
	application("npc_sxr")=ysjh_name
end if
if application("npc_sxr")=ysjh_name then
application("npc_sxsj")=now()	'npc刷新时间
'NPC自动发放
zdnpc_ok="no"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ysjh_usermdb")
if chatinfo(0)=ysjh_chat3 and allnpc="" and Minute(time())<15 then
	Application.Lock
	rs.open "select * from [npc] order by id desc",conn,3,3
        sl=rs.recordcount
        randomize()
        npcid=int(rnd*(sl))
        rs.move npcid
        zdnpc_name=rs("n姓名")
	if Instr(Application("ysjh_npc"),";" & zdnpc_name & "|")=0 then
        zdnpc_id=rs("id")
        zdnpc_sex=rs("n性别")
        zdnpc_jhtx=rs("n图片")
        zdnpc_nccc=rs("n出场词")
        zdnpc_j=int(sqr(int(rs("n经验")/50)))
        conn.execute  "update npc set n出现率=n出现率+1,n主人='无',n敌人='无',n攻击="&zdnpc_j*100&",n复活时间='"&now()&"',n防御="&zdnpc_j*150&",n武功="&zdnpc_j*800&",n体力="&zdnpc_j*5000&",n内力="&zdnpc_j*150&",n银两="&zdnpc_j*150000&",n攻击时间='"&now()&"' where n姓名='"&zdnpc_name&"'"
        myonline =zdnpc_name& "|" &zdnpc_sex& "|NPC|"&rs("n主人")&"|" &zdnpc_jhtx& "|" &zdnpc_j& "|"&zdnpc_id&"|0|1|"&"正常"
        rs.close
        call addnpc(zdnpc_name)
        mess="<bgsound src=readonly/cd.mid loop=1><font color=black>【NPC自动进入】" & zdnpc_name &"</a></font>"&zdnpc_nccc
	call showchat(mess)
	zdnpc_ok="yes"
	else
		rs.close
	end if
	Application.UnLock
end if
'NPC自动发放
if chatinfo(0)=ysjh_chat3 and zdnpc_ok="no" then
 npc_room=split(Application("ysjh_npc"),"|")
 npc_inroom=ubound(npc_room)
 if npc_inroom>=1 then
    randomize timer
    npcr=int(rnd*20)+1
    mekill=50*npcr
    nl=int(rnd*npc_inroom)
    nl1=npc_room(nl)
    nl2=split(nl1,";")
    nl_name=trim(nl2(1))
    select case npcr
    case 1,5,8,18
	npcr2=npcr
	sql="select n主人 from npc where n姓名='"&nl_name&"'"
	rs.open sql,conn,2,2
	if not rs.eof then
		n_zhuren=rs("n主人")
	end if
	rs.close
	if n_zhuren<>ysjh_name then
	 sql="select * from 用户 where 姓名='"&ysjh_name&"'"
	 rs.open sql,conn,2,2
	 if rs("保护")=false and rs("门派")<>"出家" and rs("等级")>18 then
		select case npcr2
		case 1
		mess="<font color=red>【NPC攻击】</font>["&nl_name&"]用[NPC家族]招式{绝对诱惑}攻击"&ysjh_name&",杀伤" &  ysjh_name &"体力:<font color=red>-"& mekill&"</font>点，" &  ysjh_name &"气的直发晕……"
		conn.execute "update 用户 set 体力=体力-"& mekill &" where 姓名='" &  ysjh_name &"'"
		conn.execute  "update npc set n攻击率=n攻击率+1 where n姓名='"&nl_name&"'"
		case 5
		mess="<font color=red>【NPC必杀技】</font>["&nl_name&"]用[NPC家族]招式{一掌四式}攻击"&ysjh_name &",杀伤" &  ysjh_name &"体力:<font color=red>-"& mekill&"</font>点，打的" &  ysjh_name &"晕头转向……"
		conn.execute "update 用户 set 体力=体力-"& mekill &" where 姓名='" &  ysjh_name &"'"
		conn.execute  "update npc set n攻击率=n攻击率+1 where n姓名='"&nl_name&"'"
		case 8
		mess="<font color=red>【补充体力】</font>["&nl_name&"]感觉体力不支，急忙使用了<font color=red>[NPC专用特效补药]100个</font>,体力马上暴涨了<font color=red>1000</font>点,内力暴涨<font color=red>800</font>点……"
		conn.execute  "update npc set n体力=n体力+1000,n内力=n内力+800,n攻击率=n攻击率+1 where n姓名='"&nl_name&"'"
		case 18
		mess="<font color=red>【停止攻击】</font><font color=green>["&nl_name&"]</font>觉得好困，真想休息！</font>"
		end select
		call showchat(mess)
	 end if
	 rs.close
	end if
    case 3,6,7,9,10,12,14,20
	npcr2=npcr
	select case npcr2
	case 3
		mess="<font color=red>【NPC生气】</font>NPC人物<font color=blue>["&nl_name&"]</font>觉得在这个世界上,没有一个人关心他,越想越气!~~"
	case 6
		mess="<font color=red>【NPC感慨】</font>NPC人物<font color=blue>["&nl_name&"]</font>自言自语的说：“想我这一生纵横江湖,却敌不过["&ysjh_name&"]这个小辈!”"
	case 7
		mess="<font color=red>【NPC叫嚷】</font>NPC人物<font color=blue>["&nl_name&"]</font>看见大家都无精打采的,大叫道:下雨拉,收衣服拉!~~"
	case 9
		mess="<font color=red>【NPC闲逛】</font>NPC人物<font color=blue>["&nl_name&"]</font>在江湖逛了一圈,没有找到一个称心的猎物"
	case 10
		mess="<font color=red>【NPC唱歌】</font>NPC人物<font color=blue>["&nl_name&"]</font>实在太无聊了,唱到:妹妹你做船头....."
	case 12
		mess="<font color=red>【NPC挑拨】</font><font color=blue>["&nl_name&"]</font>指着["&ysjh_name&"]骂道：“你这个小人,你开了保护和我打呀,不然你就一辈子当和尚出家,要不就穿着你的马甲保护好！”"
	case 14
		mess="<font color=red>【NPC动作】</font>NPC人物<font color=blue>["&nl_name&"]</font>感觉到腿有点酸,看看天,恩,明天要下雨了....</font>"
	case 20
		mess="<font color=red>【NPC恶作剧】</font>NPC人物<font color=blue>["&nl_name&"]</font>敲了敲["&ysjh_name&"]的脑袋，咚咚作响，打开一看，全是乱麻似的电线，不禁摇了摇头道：“这种机器人也太原始了！”"
	end select
	call showchat(mess)
    end select
 end if
end if
end if
'npc自动攻击
randomize()
rnd1=int(rnd*400)+1
says=""
kill=int(rnd*1000)+400
mekill=kill*10
npc_room=split(Application("ysjh_npc"),"|")
npc_inroom=ubound(npc_room)
if npc_inroom>=1 then
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("ysjh_usermdb")
randomize timer
nl=int(rnd*npc_inroom)
nl1=npc_room(nl)
nl2=split(nl1,";")
nl_name=trim(nl2(1))
sql="select * from npc where n姓名='"&nl_name&"'"
rs.open sql,conn,2,2
if not rs.eof then
    n_diren=rs("n敌人")
end if
rs.close
if n_diren=ysjh_name then
sql="select * from 用户 where 姓名='"&ysjh_name&"'"
rs.open sql,conn,2,2
if rs("保护")=true and rs("门派")<>"出家" and rs("等级")>21 then
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动用卡】<font color=balck>"&ysjh_name&"以为自己练功保护了就没事,<img src=sjfunc/1.gif>哪曾想NPC["&nl_name&"]使出一张解除卡，轻松的解除了"&ysjh_name&"的保护状态,这就是成为NPC["&nl_name&"]敌人的下场,保护了也无用!</font>"
call showchat(says)
conn.execute  "update 用户 set 保护=false where 姓名='"&ysjh_name &"'"
end if
if rs("保护")=false and rs("门派")<>"出家" and rs("等级")>21 then
rs.close
dz=int(rnd*80)+1
sql="select * from y where id="&dz&""
rs.open sql,conn,2,2
 dz_ok=rs("a")
rs.close
npcdj=int(rnd*5)+1
Select Case npcdj
case 1
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动攻击】<font color=balck>"&ysjh_name&"正在自鸣得意<img src=sjfunc/1.gif>突然被NPC["&nl_name&"]使出一招["&dz_ok&"]，伤得不轻，体力下降"&mekill&"点，内力下降"&kill&"点，这就是成为NPC["&nl_name&"]敌人的下场!</font>"
case 2
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动攻击】<font color=balck>"&ysjh_name&"正自以为是<img src=sjfunc/1.gif>哪里料到NPC["&nl_name&"]突然使出一招["&dz_ok&"]，伤得不轻，体力下降"&mekill&"点，内力下降"&kill&"点，这就是成为NPC["&nl_name&"]敌人的下场!</font>"
case 3
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动攻击】<font color=balck>"&ysjh_name&"正在家中休息<img src=sjfunc/1.gif>却被NPC["&nl_name&"]使出一招["&dz_ok&"]，伤得不轻，体力下降"&mekill&"点，内力下降"&kill&"点，这就是成为NPC["&nl_name&"]敌人的下场!</font>"
case 4
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动攻击】<font color=balck>"&ysjh_name&"正在花园赏花<img src=sjfunc/1.gif>突然NPC["&nl_name&"]从花丛中蹦出,使出一招["&dz_ok&"]，伤得不轻，体力下降"&mekill&"点，内力下降"&kill&"点，这就是成为NPC["&nl_name&"]敌人的下场!</font>"
case 5
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>【NPC自动攻击】<font color=balck>"&ysjh_name&"正在同情人风花雪月<img src=sjfunc/1.gif>几曾想被NPC["&nl_name&"]使出一招["&dz_ok&"]，伤得不轻，体力下降"&mekill&"点，内力下降"&kill&"点，这就是成为NPC["&nl_name&"]敌人的下场!</font>"	
end select
call showchat(says)
conn.execute  "update 用户 set 体力=体力-"&mekill&",内力=内力-"&kill&" where 姓名='"&ysjh_name &"'"
conn.execute  "update npc set n攻击率=n攻击率+1 where n姓名='"&nl_name&"'" 
end if
end if
end if
'npc自动攻击
'NPC自动退出
if Application("ysjh_npc")<>"" and chatinfo(0)=ysjh_chat3 then
  npc_room=split(Application("ysjh_npc"),"|")
  npc_inroom=ubound(npc_room)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("ysjh_usermdb")
  for i=0 to npc_inroom-1
	nl1=npc_room(nl)
	nl2=split(nl1,";")
	nl_name=trim(nl2(1))
        sql="select * from [npc] where n姓名='"&nl_name&"'"
        rs.open sql,conn,1,1
        sj=DateDiff("n",rs("n攻击时间"),now())
        if sj>15 then
	call delnpc(nl_name)
	mess= "<bgsound src=wav/gf.wav loop=1><font color=black>【NPC自动退出】</font>NPC人物["&nl_name&"]晃来晃去，没找到好吃的东西，跑了几圈江湖，渐渐地消失了..."
	call showchat(mess)
        end if
        rs.close
  next
end if
'NPC自动退出
%>