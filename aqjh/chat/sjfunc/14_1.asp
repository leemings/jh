<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="../../const3.asp"-->
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
onlinekill=aqjh_onlinekill
if onlinenow<onlinekill and chatinfo(0)<>aqjh_chat1 and chatinfo(0)<>aqjh_chat2 then
	Response.Write "<script language=JavaScript>{alert('提示:非恩怨房在线低于"&onlinekill&"人不得动武！');}</script>"
	Response.End
end if
next
%>
<%'发招(比拼内力)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_fz then
	Response.Write "<script language=JavaScript>{alert('提示：发招需要["&jhdj_fz&"]级以上才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
'判断对npc操作
if InStr(";" & Application("aqjh_npc"), ";" & towho & "|")<>0 then
   pk="【NPC攻击】"
else
   pk="【动武】"
end if
'判断对npc操作
says="<font color=red>"&pk&"<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs.eof then
   rs.close
   rs.open "select * from [npc] where n姓名='" & to1 &"'",conn,2,2
   if rs.eof then
      wt=""
   else
      wt="npc"
   end if
else
   wt="ren"
end if
rs.close
select case wt
case "npc"
'开始判断npc
rs.open "select * from npc where n姓名='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：NPC["&to1&"]并不存在或者它不是NPC你搞错了吧！');}</script>"
	Response.End	
end if
if rs("n体力")<0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
'删除在线NPC名单
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),to1 & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & to1 & " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
attack="##:NPC已经无力反抗了，官府已经收回！"
exit function
end if
rs.close
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("保护")=True and aqjh_grade<10 then
    attack="##:你正在保护中，请不要对NPC进行攻击！"
    exit function
end if
if rs("门派")="出家" then
    attack="##:你是出家人，不能参与NPC打架！"
    exit function
end if
yxjfyfrom=rs("防御")
yxjgjfrom=rs("攻击")
yxjwgfrom=rs("武功")
yxjdjfrom=rs("等级")
nlfrom=rs("内力")
menpai=rs("门派")
rs.close
'开始武功
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
if rs.eof or rs.bof then
    attack="##:你的门派:"&menpai&"并没有这样的武功["&fn1&"]呀！"
    exit function
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
wgsm=rs("f")
if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*130)+1 &".gif height=70>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&" height=70>"
end if
if rs("c")>nlfrom or rs("d")>yxjwgfrom then
    attack="##:您的内力或武功不足，不能对npc进行攻击！先补一补吧！"
    exit function
end if
rs.close
'对NPC的判断
rs.open "select * from npc where n姓名='" & to1 &"'",conn
to1=rs("n姓名")
yxjgjto=rs("n攻击")
yxjfyto=rs("n防御")
yxjwgto=rs("n武功")
yxjdjto=int(sqr(int(rs("n经验")/50)))
nlto=rs("n内力")
nzhuren=rs("n主人")
ndiren=rs("n敌人")
sj=DateDiff("n",rs("n复活时间"),now())
rs.close
'判断主人
if nzhuren=aqjh_name then
    attack="##:NPC%%是你的部下，你不能对它攻击，不怕它背叛你吗？"
    exit function
end if
'判断敌人
if sj>5 then
if ndiren<>aqjh_name and ndiren<>"无" and Instr(LCase(useronlinename),LCase(" "&ndiren&" "))<>0 then
    attack="##:NPC%%正和"&ndiren&"打架中，请不要插手，不然%%咬死你！"
    exit function
end if
end if
randomize timer
ki=int(rnd*6)+4
killer=int(((yxjgjfrom+yxjwgfrom+yxjfyfrom)-(yxjfyto+yxjgjto+yxjwgto)*(yxjdjfrom/yxjdjto))/ki)+wgwg+wgnl+wgdj*100
if killer>0 then
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& int(killer/30) & " where 姓名='" & aqjh_name &"'"
	conn.execute "update npc set n体力=n体力-"&int(killer/15)&",n敌人='"&aqjh_name&"',n复活时间=now(),n攻击时间=now() where n姓名='" & to1 &"'"
	djinfo="##武功高强"
        mekill=abs(int(killer/30))
        killer=int(killer/15)
else
	if abs(killer)>5000 then killer=5000
        killer=int(abs(killer)/10)
	mekill=killer*10
        conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"&mekill&" where 姓名='" & aqjh_name &"'"
	conn.execute "update npc set n体力=n体力-"&killer&",n敌人='"&aqjh_name&"',n复活时间=now(),n攻击时间=now() where n姓名='" & to1 &"'"
	djinfo="##武功技不如人"
end if
rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*23)+1
mywav="wav"&wav&".wav"
randomize timer
dz=int(rnd*9)+1
select case dz
case 1
 dz_ok="天下无狗"
case 2
 dz_ok="拨狗朝天"
case 3
 dz_ok="棒打双犬"
case 4
 dz_ok="乾坤借法"
case 5
 dz_ok="昆山叠掌"
case 6
 dz_ok="金蛇吐雾"
case 7
 dz_ok="天外飞龙"
case 8
 dz_ok="破玉拳"
case 9
 dz_ok="神剑伏魔"
end select
if rs("体力")<-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得NPC[%%]体力下降:<font color=red>-"& killer &"</font>点,自己被怪兽[%%]有了一招<fonr color=red>["&dz_ok&"]</font>,使得体力下降:<font color=red>-"& mekill &"</font>点,##武技不如人，被%%活活咬死……"
        conn.execute "update npc set n敌人='无' where n姓名='" & to1 &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        call boot(aqjh_name,"被NPC咬死，操作者："&aqjh_name&",["&menpai&"]唉，技不如人"&fn1)
	exit function
end if
rs.close
rs.open "select * from npc where n姓名='" & to1 &"'",conn,2,2
zlwp=rs("n物品")
zlyl=rs("n银两")
n_jy=rs("n经验")
n_dj=int(sqr(int(rs("n经验")/50)))
n_baochu=(n_dj/aqjh_jhdj)*1000
n_jjy=int((n_dj/aqjh_jhdj)*500)
add_type=rs("n物品类型")
nzhuren=rs("n主人")
if zlyl>0 then
   njy="<font color=red>【NPC加血】<font color=" & saycolor & ">NPC[%%]花了"&zlyl&"银白银，补足了体力<font color=red>"&zlyl*0.1&"</font>点。<br>"
   conn.execute "update npc set n银两=0,n体力=n体力+"&zlyl*0.1&" where n姓名='" & to1 &"'"
end if
if rs("n体力")>-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得NPC[%%]体力下降:<font color=red>-"& killer &"</font>点。<br>"
        attack=attack&njy&"<font color=red>【NPC反击】<font color=" & saycolor & ">由于##惹到了NPC[%%]，NPC[%%]奋一切力量，使出了一招<font color=red>["&dz_ok&"]</font>，使得##体力下降:<font color=red>-"& mekill &"</font>点."
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
'是否有暴出物品
if n_baochu>=Application("aqjh_npcwp") then
'暴出物品
data1=split(zlwp,";")
data2=ubound(data1)
for kk=0 to data2-1
    data3=split(data1(kk),"|")
    wp_sl=data3(1)
    wp_name=trim(data3(0))
    Application.Lock
    Application("npc_wp"&kk)=1
    Application.UnLock
    randomize()
    regjm=int(rnd*3348998)
    wp_list=wp_list&"<input type=button value='"&wp_name&"|"&wp_sl&"个' name=npc"&regjm&" style='background-color:#86A231;color:FFFFFF;border: 1 double' onClick=javascript:npc"&regjm&".disabled=1;window.open('npc_wp.asp?dj="&aqjh_jhdj&"&wp_name="&wp_name&"&wp_sl="&wp_sl&"&kk="&kk&"&wp_type="&add_type&"','d','scrollbars=0,resizable=0,width=450,height=500')> "
next
    '是否有特殊装备
     if aqjh_jhdj<n_dj then
       rs.close
       rs.open "SELECT * FROM x where f='极品' order by g",conn,1,2
       wp_count=rs.recordcount+2
       randomize timer
       wp_rnd=int(rnd*wp_count)+1
       if wp_rnd<rs.recordcount then
          rs.moveFirst
          rs.move wp_rnd
          wp_jian=rs("a")
       else
          wp_jian="无"
       end if
       if wp_jian<>"无" then
        rs.close
        rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
        zstemp=add(rs("w3"),wp_jian,1)
        conn.execute "update 用户 set w3='"&zstemp&"' where  姓名='"&aqjh_name&"'"
        jian="<font color=red><b>##得到一把"&wp_jian&"！</b></font>"
       else
        jian="<font color=red><b>##只见到一把不寻常的宝剑瞬间消失了，唉，只怪自己运气不好！</b></font>"
       end if
     end if 
conn.execute "update npc set n敌人='"&aqjh_name&"',n死次数=n死次数+1,n死亡时间='"&now()&"' where n姓名='" & to1 &"'"
attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],%%哪是##的对手，已经奄奄一息.."&jian&"[%%]身上所有物品四分五裂，撒落了一地，大家快抢呀！"&wp_list
'战利品
else
  attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],%%哪是##的对手，已经奄奄一息...但是，没得到任何东西！"
end if
conn.execute "update 用户 set allvalue=allvalue+"&n_jjy&" where 姓名='" &aqjh_name&"'"
attack=attack&"<br><font color=ff0099>【恭喜】</font>##打死了NPC人物%%，经验值上升<font color=red>"&n_jjy&"</font>点。"
'删除在线NPC名单
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),to1 & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & to1 & " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
case "ren"
rs.open "select 转生,状态,等级 from 用户 where 姓名='" & aqjh_name &"'",conn
zstt=rs("转生")
todj=rs("等级")
rs.close
rs.open "select 门派,等级,保护,grade,宝物,杀人时间,登录,转生 from 用户 where 姓名='" & to1 &"'",conn,2,2
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，%%上线还没到3分钟啊！"
        exit function
end if
dlsj=DateDiff("n",rs("杀人时间"),now())
if dlsj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，<font color=red><b>你不能对%%进行攻击，他刚刚被人杀死，请等30分钟后吧！</b></font>"
        exit function
end if
if rs("等级")<=90 and rs("转生")=0 and todj>=90 then
if rs("等级")+100<=todj  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##对%%动武失败，原因：%%不满90级,且##级别比%%高100以上,不能动武啊。</b></font>"
	exit function
end if
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，他是出家人！"
        exit function
end if
if rs("门派")="官府" and aqjh_grade<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对官府的[%%]进行攻击！"
        exit function
end if
if rs("等级")<=30 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，不要对初入江湖新手操作！"
        exit function
end if
if rs("保护")=True and aqjh_grade<10 and zstt<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，对方正在练功保护请不要偷袭!"
        exit function
end if
rs.close
rs.open "select 门派,保护,防御,转生,等级,攻击,武功,内力,门派,杀人数,暴豆时间,grade,死亡时间,登录,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你是出家人不可以操作!"
        exit function
end if
if rs("等级")>90 and zstt=0 and todj<90 then
if rs("等级")-100>todj then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##对%%动武失败，原因：%%不满90级,且##级别比%%高100以上,不能动武哦。</b></font>"
	exit function
end if
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你上线还不到3分钟就不要再杀人了！"
        exit function
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你正在练功保护请不要发招！"
        exit function
end if
if rs("杀人数")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再发招了！"
        exit function
end if
sj=DateDiff("n",rs("暴豆时间"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=red><b>威力爆满</b></font>"
	baodoudj=3
end if
yxjfyfrom=rs("防御")
yxjgjfrom=rs("攻击")
yxjwgfrom=rs("武功")
nlfrom=rs("内力")
menpai=rs("门派")
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
'取出自己装备武器的攻击及特性(并且所有武器耐久-1
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsm="##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=5 then	
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		yxjtx=data(4)
		gjtx="<font color=red><b>["&data(0)&"]</b></font>"
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
conn.execute "update 用户 set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  姓名='"&aqjh_name&"'"
rs.close
'开始武功
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的门派:"& menpai &" 并没有这样的武功["& fn1 &"]呀！');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
wgsm=rs("f")
if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*130)+1 &".gif height=70>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&" height=70>"
end if
if rs("c")>nlfrom or rs("d")>yxjwgfrom then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您的内力或武功不足!');}</script>"
	Response.End
end if
'判断对方有没有npc
rs.close
rs.open "select top 1 * from NPC where n主人='"&to1&"'",conn
if not rs.eof then
  n_name=rs("n姓名")
  n_tl=rs("n体力")
  n_nl=rs("n内力")
  n_wg=rs("n武功")
  n_gj=rs("n攻击")
  n_fy=rs("n防御")
  n_diren=rs("n敌人")
  if InStr(";" &Application("aqjh_npc"), ";" & n_name & "|")<>0 then
  isnpc=true
  end if
end if
'判断对方有没有npc
'判断自己有没有npc
rs.close
rs.open "select top 1 * from NPC where n主人='"&aqjh_name&"'",conn
if not rs.eof then
  my_n_name=rs("n姓名")
  my_n_dj=int(sqr(int(rs("n经验")/50)))
  my_n_diren=rs("n敌人")
  if InStr(";" &Application("aqjh_npc"), ";" & my_n_name & "|")<>0 then
  ismynpc=true
  end if
end if
'判断自己有没有npc
'对方的判断
rs.close
rs.open "select 姓名,攻击,防御,武功,等级,会员等级,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
to1_dj=rs("等级")
yxjgjto=rs("攻击")
yxjfyto=rs("防御")
yxjwgto=rs("武功")
jhhy=rs("会员等级")
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
'取出对方装备武器的攻击及特性(并且所有武器耐久-1)
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsmto="%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		yxjtotx=data(4)
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
conn.execute "update 用户 set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  姓名='"&to1&"'"
rs.close
randomize timer
ki=int(rnd*6)+4
killer=int(((yxjgjfrom+yxjwgfrom+yxjfyfrom)-(yxjfyto+yxjgjto+yxjwgto))/ki)*baodoudj+wgwg+wgnl+wgdj*100
if isnpc=true then
npc_kill=int((yxjgjfrom+yxjwgfrom+yxjfyfrom)/2-(n_gj+n_fy+n_wg))
end if
randomize timer
'对武功次数计数
if menpai<>"游侠" then
	conn.execute "update y set e=e+1 where id="&wgid
end if
randomize timer
ki=int(rnd*500)
if yxjtx<>"无" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器克制对方，吸取对方[<font color=blue><b>"&yxjtx&"</b></font>]共计:+<font color=red>"&ki&"</font>点!"
	else
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器被对方所克制，被吸取[<font color=blue><b>"&yxjtx&"</b></font>]共计:-<font color=red>"&ki&"</font>点"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& int(killer/10) & " where 姓名='" & aqjh_name &"'"
'判断npc
        if isnpc=true then
	conn.execute "update 用户 set 体力=体力-"  & int(killer/2) & " where 姓名='" & to1 &"'"
        conn.execute "update 用户 set 内力=int(内力-内力*0.07),武功=int(武功-武功*0.07),体力=int(体力-体力*0.1) where 姓名='" &aqjh_name &"'"
        npc_attack="<font color=red>【NPC保护主人攻击】</font>由于##对%%进行攻击,NPC人物"&n_name&":护主攻击!但由于##武功高强,只受了点轻伤！内力，武功各下降了1/14，体力下降了1/10..."
        else
	conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 &"'"
        end if
	djinfo="##武功高强"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
'判断npc
        if isnpc=true then
        npc_kill=abs(npc_kill)
        conn.execute "update 用户 set 内力=内力-"&int(npc_kill/20)& ",武功=武功-"&int(npc_kill/20)&",体力=体力-"&int(npc_kill/10)&" where 姓名='" & aqjh_name &"'"
        npc_attack="<font color=red>【NPC保护主人攻击】</font>由于##对%%进行攻击,NPC人物"&n_name&"护主攻击!但由于##武功技不如人,内力、武功各下降<font color=red>-"&int(npc_kill/20)&"</font>点，体力下降<font color=red>-"&int(npc_kill/10)&"</font>点。"
        end if
       	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& abs(killer) & " where 姓名='" & aqjh_name &"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
	djinfo="##武功技不如人"
end if
'npc判断
if isnpc=true then
 if Instr(LCase(useronlinename),LCase(" "&n_diren&" "))=0 or n_diren="无" then
       conn.execute "update npc set n敌人='"&aqjh_name&"',n攻击时间=now() where n姓名='"&n_name&"'"
 end if
end if
if ismynpc=true then
 if Instr(LCase(useronlinename),LCase(" "&my_n_diren&" "))=0 or my_n_diren="无" then
     conn.execute "update npc set n敌人='"&to1&"',n攻击时间=now() where n姓名='"&my_n_name&"'"
 end if
 my_abate=int(my_n_dj*500/to1_dj)
 my_npc_attack="<font color=red>【NPC参战出击】</font>NPC人物["&my_n_name&"]看到主人攻击%%,上去就咬了%%一口,%%损失<font color=red>体力-"&my_abate&"</font>"
 conn.execute "update 用户 set 体力=体力-"&my_abate&" where 姓名='"&to1&"'"
end if
rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
if rs("体力")<-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点,由于%%的内力深厚，##只感觉眼前一些，双腿一蹬，灵魂飞往封神台报到去……"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(aqjh_name,"发招，操作者："&aqjh_name&",["&menpai&"]唉，技不如人"&fn1)
	exit function
end if
rs.close
rs.open "select 体力,宝物,门派,通缉 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("体力")>-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx&"<br>"&my_npc_attack&"<br>"&npc_attack
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name &"'"
	if rs("宝物")=Application("aqjh_baowuname") then
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set 宝物修练=0,宝物='无',保护=true,内力=100,体力=2000 where 姓名='" & to1 &"'"
		attack="##把%%的宝物:"& Application("aqjh_baowuname") &"抢走，因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&gjtx
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		exit function
	end if
if rs("通缉")=False then
	conn.execute "update 用户 set 杀人时间=now(),状态='死',事件原因='"&aqjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
	attack=xinxi & "##<bgsound src=wav/daipu.wav loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,%%慢慢倒下..江湖从此又少了一位大虾!"&gjtx&"<br>"&npc_attack
else
	conn.execute "update 用户 set 通缉=False,状态='死',道德=0,体力=0,内力=0,武功=0,魅力=0,银两=0,存款=int(存款/2),事件原因='"&aqjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
	conn.execute "update 用户 set 银两=银两+5000000 where 姓名='" & aqjh_name &"'"
	attack="攻击通缉犯，"&xinxi & "##<bgsound src=wav/daipu.wav loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,%%慢慢倒下..为此：##得到500万元官府奖励!"&gjtx&"<br>"&npc_attack
end if
call boot(to1,"发招，操作者："&aqjh_name&",["&menpai&"]"&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','"&fn1&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
case else
 Response.Write "<script language=JavaScript>{alert('提示：找不到攻击对象！');}</script>"
 Response.End
end select
end function
%>