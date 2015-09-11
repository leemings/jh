<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
onlinekill=Application("aqjh_onlinekill")
if onlinenow<onlinekill and chatinfo(0)<>"决战江湖" then
	Response.Write "<script language=JavaScript>{alert('提示:大厅在线低于"&onlinekill&"人不得动武！');}</script>"
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
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
f=Minute(time()) 
if f>Application("aqjh_pktime") and chatinfo(0)<>"决战江湖" then
        Response.Write "<script language=JavaScript>{alert('提示：江湖PK打架时间为每小时前["&Application("aqjh_pktime")&"]分，打架请去恩怨！');}</script>"
	response.end
end if
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
Application("aqjh_npc"&nowinroom)=Replace(Application("aqjh_npc"&nowinroom),";"&to1&"|","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
attack="##:NPC已经无力反抗了，官府已经收回！"
exit function
end if
rs.close
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
zhaoshi=trim(rs("招式") & " ")
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
fn1=trim(fn1)
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='NPC'",session("bug_conn_npc")
if instr("," & zhaoshi , "," & fn1 & "|")=0 then response.write "<script language=JavaScript>{alert('你不会这种武功"&zhaoshi &"');}</script>" : response.end
if rs.eof then
	rs.close
	set rs=nothing	
	Response.Write "<script language=JavaScript>{alert('提示：没有这样的武功["& fn1 &"]呀！');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
wgsm=rs("f")
if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*98)+1 &".gif>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&">"
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
	conn.execute "update npc set n体力=n体力-"&int(killer/15)&",n敌人='"&aqjh_name&"',n复活时间=now() where n姓名='" & to1 &"'"
	djinfo="##武功高强"
        mekill=abs(int(killer/30))
        killer=int(killer/15)
else
	if abs(killer)>5000 then killer=5000
        killer=int(abs(killer)/10)
	mekill=killer*10
        conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"&mekill&" where 姓名='" & aqjh_name &"'"
	conn.execute "update npc set n体力=n体力-"&killer&",n敌人='"&aqjh_name&"',n复活时间=now() where n姓名='" & to1 &"'"
	djinfo="##武功技不如人"
end if
rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
randomize timer
dz=int(rnd*9)+1
select case dz
case 1
 dz_ok="降龙十八掌"
case 2
 dz_ok="拨草寻蛇"
case 3
 dz_ok="药师佛根本印"
case 4
 dz_ok="剑花烟雨江南"
case 5
 dz_ok="太极神拳"
case 6
 dz_ok="金蛇吐信"
case 7
 dz_ok="天外飞仙"
case 8
 dz_ok="八仙过海"
case 9
 dz_ok="天涯明月刀"
end select
if rs("体力")<-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>用个人修炼招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得NPC[%%]体力下降:<font color=red>-"& killer &"</font>点,自己被怪兽[%%]有了一招<fonr color=red>["&dz_ok&"]</font>,使得体力下降:<font color=red>-"& mekill &"</font>点,##武技不如人，被%%活活咬死……"
        conn.execute "update npc set n敌人='无' where n姓名='" & to1 &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        call boot(aqjh_name,"被NPC打死，操作者："&aqjh_name&",["&menpai&"]唉，技不如人"&fn1)
	exit function
end if
rs.close
rs.open "select * from npc where n姓名='" & to1 &"'",conn,2,2
zlwp=rs("n物品")
zlyl=rs("n银两")
n_jy=rs("n经验")
n_dj=int(sqr(int(rs("n经验")/50)))
n_baochu=(n_dj/aqjh_jhdj)*1000
wplx=rs("n物品类型")
       select case wplx
       case "药品"
         add_type="w1"
       case "暗器"
         add_type="w4"
       case "毒药"
         add_type="w2"
       case "鲜花"
         add_type="w7"
       case "法器"
         add_type="w10"
       case "书籍"
         add_type="w9"
       case "卡片"
         add_type="w5"
       end select
nzhuren=rs("n主人")
if zlyl>0 then
   njy="<font color=red>【NPC补体力】<font color=" & saycolor & ">NPC[%%]花了"&zlyl&"银白银，补足了体力<font color=red>"&zlyl*0.01&"</font>点。<br>"
   conn.execute "update npc set n银两=0,n体力=n体力+"&zlyl*0.01&" where n姓名='" & to1 &"'"
end if
if rs("n体力")>-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>用用个人修炼招式招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得NPC[%%]体力下降:<font color=red>-"& killer &"</font>点。<br>"
        attack=attack&njy&"<font color=red>【NPC攻击】<font color=" & saycolor & ">由于##惹到了NPC[%%]，NPC[%%]奋一切力量，使出了一招<font color=red>["&dz_ok&"]</font>，使得##体力下降:<font color=red>-"& mekill &"</font>点."
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
'是否有战利品
if n_baochu>=Application("aqjh_nptfp") then
'战利品

conn.execute "update npc set n敌人='"&aqjh_name&"',n死次数=n死次数+1,n死亡时间='"&now()&"' where n姓名='" & to1 &"'"
s=int(rnd*8898)+10000
	s1=int(rnd*8898)+1000
	s2=int(rnd*9898)+100
	s3=int(rnd*9898)+10
	s4=int(rnd*9898)+50
	s5=int(rnd*9898)+500
	s6=int(rnd*9898)+5000
	s7=int(rnd*9898)+5000
	s8=int(rnd*9898)+5001
        s9=int(rnd*9898)+5000
	Application.Lock
	Application("aqjh_zlp1")=s
	Application("aqjh_zlp2")=s1
	Application("aqjh_zlp3")=s2
	Application("aqjh_kp3")=s3
	Application("aqjh_zlp5")=s4
	Application("aqjh_zlp6")=s5
	Application("aqjh_zlp7")=s6
	Application("aqjh_zlp8")=s7
	Application("aqjh_zlp9")=s8
        Application("aqjh_zlp10")=s9
	Application.UnLock

randomize timer
npcrr=int(rnd*7)+1
select case npcrr
case 1
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='银两'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('ylok.asp?tl4="&Application("aqjh_zlp9")&"','d') name=tongyi"&regjm&"><input type=button value='配药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('pyokl.asp?tl3="&Application("aqjh_zlp2")&"','d') name=tongyi"&regjm&"><input type=button value='物品'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('hrokl.asp?tl4="&Application("aqjh_zlp1")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊......"
case 2
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='药品'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('ypokl.asp?tl1="&Application("aqjh_zlp3")&"','d') name=tongyi"&regjm&"><input type=button value='鲜花'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('huaok.asp?tl4="&Application("aqjh_zlp5")&"','d') name=tongyi"&regjm&"><input type=button value='银两'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('ylok.asp?tl4="&Application("aqjh_zlp9")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊......"
case 3
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%], 可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='银两'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('ylok.asp?tl4="&Application("aqjh_zlp9")&"','d') name=tongyi"&regjm&"><input type=button value='暗器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('anqiok.asp?tl4="&Application("aqjh_zlp6")&"','d') name=tongyi"&regjm&"><input type=button value='物品'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('hrokl.asp?tl4="&Application("aqjh_zlp1")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊......"
case 4
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='配药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('pyokl.asp?tl3="&Application("aqjh_zlp2")&"','d') name=tongyi"&regjm&"><input type=button value='暗器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('anqiok.asp?tl4="&Application("aqjh_zlp6")&"','d') name=tongyi"&regjm&"><input type=button value='装备'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('bqok.asp?tl4="&Application("aqjh_zlp8")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊......"
case 5
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='毒药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('dyok.asp?tl4="&Application("aqjh_zlp7")&"','d') name=tongyi"&regjm&"><input type=button value='药品'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('ypokl.asp?tl1="&Application("aqjh_zlp3")&"','d') name=tongyi"&regjm&"><input type=button value='暗器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('anqiok.asp?tl4="&Application("aqjh_zlp6")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊，......"
case 6
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='武器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('bqok.asp?tl4="&Application("aqjh_zlp8")&"','d') name=tongyi"&regjm&"><input type=button value='鲜花'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('huaok.asp?tl4="&Application("aqjh_zlp5")&"','d') name=tongyi"&regjm&"><input type=button value='毒药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('dyok.asp?tl4="&Application("aqjh_zlp7")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&"><input type=button value='书籍'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('xuok.asp?tl4="&Application("aqjh_zlp10")&"','d') name=tongyi"&regjm&">大家快抢啊，......"
case 7
	attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai&"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],  可怜的%%技不如人，被##活活的打死后，暴出了很多好东西：<input type=button value='武器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('bqok.asp?tl4="&Application("aqjh_zlp8")&"','d') name=tongyi"&regjm&"><input type=button value='鲜花'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('huaok.asp?tl4="&Application("aqjh_zlp5")&"','d') name=tongyi"&regjm&"><input type=button value='毒药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('dyok.asp?tl4="&Application("aqjh_zlp7")&"','d') name=tongyi"&regjm&"><input type=button value='暗器'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('anqiok.asp?tl4="&Application("aqjh_zlp6")&"','d') name=tongyi"&regjm&"><input type=button value='配药'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('pyokl.asp?tl3="&Application("aqjh_zlp2")&"','d') name=tongyi"&regjm&"><input type=button value='物品'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('hrokl.asp?tl4="&Application("aqjh_zlp1")&"','d') name=tongyi"&regjm&"><input type=button value='卡片'  onClick=javascript:tongyi"&regjm&".disabled=1;tongyi"&regjm&".disabled=1;window.open('kp3.asp?tl="&Application("aqjh_kp3")&"','d') name=tongyi"&regjm&">大家快抢啊......"
end select

'战利品
else
  attack="##<bgsound src=wav/daipu.wav loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"NPC[%%],%%哪是##的对手，转眼间便魂归西土...但是，没有暴出任何物品！"
end if
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
Application("aqjh_npc"&nowinroom)=Replace(Application("aqjh_npc"&nowinroom),";"&to1&"|","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
case "ren"
rs.open "select 门派,等级,保护,grade,宝物,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，他是出家人！"
        exit function
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<15 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，他刚刚被别人杀死，还是先放过他吧！"
        exit function
end if
if rs("等级")<=18 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，不要对初入江湖新手操作！"
        exit function
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，对方正在练功保护请不要偷袭!"
        exit function
end if
rs.close
rs.open "select 门派,保护,防御,攻击,武功,内力,门派,杀人数,暴豆时间,grade,死亡时间,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对%%进行攻击，你是出家人不可以操作!"
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
	if ubound(data)=4 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z1=null
			wpsm="##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=4 then	
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z2=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=4 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z3=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=4 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z4=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=4 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		yxjtx=data(4)
		gjtx="<font color=red><b>["&data(0)&"]</b></font>"
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z5=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=4 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
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
	gif="<img src=../jhmp/gif/wg"& int(rnd*98)+1 &".gif>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&">"
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
  if InStr(";" &Application("aqjh_npc"&nowinroom), ";" & n_name & "|")<>0 then
  isnpc=true
  end if
end if
'判断对方有没有npc
'对方的判断
rs.close
rs.open "select 姓名,攻击,防御,武功,会员等级,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
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
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z1=null
			wpsmto="%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z2=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z3=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z4=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		yxjtotx=data(4)
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
		else
			z5=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=4 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)
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
       conn.execute "update npc set n敌人='"&aqjh_name&"' where n姓名='"&n_name&"'"
 end if
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
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx&"<br>"&npc_attack
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
	conn.execute "update 用户 set 状态='死',事件原因='"&aqjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
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