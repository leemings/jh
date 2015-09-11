<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('提示：您没有登陆或已经超时断开连接，请重新登陆！');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"mhmj/syl.asp")=0 then 
Response.Write "<script language=javascript>{alert('提交数据非法，程序拒绝执行！');parent.history.go(-1);}</script>" 
Response.End 
end if
if Minute(time())<40 or Minute(time())>=60 then
Response.Write "<script Language=Javascript>{alert('提示：梦幻魔界只有每个小时的后20分钟才会开！');parent.history.go(-1);}</script>" 
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="恶魔之王"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
tl=rs("体力")
nl=rs("内力")
fl=rs("法力")
dj=rs("等级")
hydj=rs("会员等级")
gj=rs("攻击")
fy=rs("防御")
wg=rs("武功")
allss=int(gj+fy+wg)
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
	case 6
		jgsj=60
end select
sj=DateDiff("s",rs("操作时间"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是"&hydj&"级会员，"&jgsj&"秒可进行一次，请等"&s&"秒后再来猎场！');parent.history.go(-1);}</script>" 
	Response.End
end if
if tl<0 then 
	Response.Write "<script Language=Javascript>alert('提示：大侠。您已经壮烈了。。下次注意了！');location.href = '../../../exit.asp';</script>"
	Response.End
end if
if nl<200000 then 
	Response.Write "<script Language=Javascript>alert('提示：内力小于200000不能进入！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if fl<110000 then 
	Response.Write "<script Language=Javascript>alert('提示：法力小于110000不能进入！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
dim js(32)
js(0) ="骨头"
js(1) ="绿宝石"
js(2) ="水晶石"
js(3) ="虎肉"
js(4) ="蓝宝石"
js(5) ="木炭"
js(6) ="硝酸"
js(7) ="硫磺"
js(8) ="红宝石"
js(9) ="僵尸牙"
js(10) ="木炭"
js(11) ="木炭"
js(12) ="僵尸血"
js(13) ="绿宝石"
js(14) ="绿宝石"
js(15) ="铁块"
js(16) ="骨头"
js(17) ="僵尸血"
js(18) ="红宝石"
js(19) ="僵尸血"
js(20) ="蓝宝石"
js(21) ="木炭"
js(22) ="绿宝石"
js(23) ="硫磺"
js(24) ="硫磺"
js(25) ="蓝宝石"
js(26) ="木炭"
js(27) ="木炭"
js(28) ="僵尸牙"
js(29) ="铁块"
js(30) ="僵尸牙"
js(31) ="铁块"
randomize()
myxy = Int(Rnd*32)
s1=Int(Rnd*20000)
s2=Int(Rnd*10000)
s3=Int(Rnd*150)
s4=Int(Rnd*4)
s5=Int(Rnd*2)
s6=Int(Rnd*300)
zstemp=add(rs("w6"),js(myxy),7)
dim js1(24)
js1(0) ="精制当归"
js1(1) ="金创药"
js1(2) ="白饭"
js1(3) ="大白菜"
js1(4) ="黄酒"
js1(5) ="大补丸"
js1(6) ="金创药"
js1(7) ="金创药"
js1(8) ="仙芝草"
js1(9) ="黄酒"
js1(10) ="白饭"
js1(11) ="黄酒"
js1(12) ="精制当归"
js1(13) ="金创药"
js1(14) ="金创药"
js1(15) ="黄酒"
js1(16) ="大白菜"
js1(17) ="大补丸"
js1(18) ="仙芝草"
js1(19) ="大白菜"
js1(20) ="金创药"
js1(21) ="枸杞子"
js1(22) ="白饭"
js1(23) ="大白菜"
randomize()
myxx = Int(Rnd*24)
zstemp1=add(rs("w1"),js1(myxx),15)
dim js2(40)
js2(0) ="太阳盾牌"
js2(1) ="月光盾牌"
js2(2) ="银盾牌"
js2(3) ="月亮刀"
js2(4) ="太阳剑"
js2(5) ="真武剑"
js2(6) ="拂柳剑"
js2(7) ="玉女剑"
js2(8) ="太阳戒指"
js2(9) ="太阳项链"
js2(10) ="月光戒指"
js2(11) ="月光项链"
js2(12) ="太阳保护"
js2(13) ="月光保护"
js2(14) ="龙靴"
js2(15) ="黄金靴"
js2(16) ="银帽"
js2(17) ="黄金帽"
js2(18) ="无情剑"
js2(19) ="胜邪剑"
js2(20) ="龙火剑"
js2(21) ="紫微剑"
js2(22) ="松纹剑"
js2(23) ="青钢剑"
js2(24) ="金蛇剑"
js2(25) ="烈焰剑"
js2(26) ="碧水剑"
js2(27) ="越女剑"
js2(28) ="君子剑"
js2(29) ="白龙剑"
js2(30) ="圣泉剑"
js2(31) ="水晶盔甲"
js2(32) ="黄金盔甲"
js2(33) ="钢铁盾牌"
js2(34) ="青铜盾牌"
js2(35) ="银盔甲"
js2(36) ="青铜甲"
js2(37) ="水晶帽"
js2(38) ="太阳帽"
js2(39) ="灵魂戒指"
randomize()
myxa = Int(Rnd*41)
zstemp2=add(rs("w3"),js2(myxa),1)
dim js3(41)
js3(0) ="补血卡"
js3(1) ="暴豆卡"
js3(2) ="练功卡"
js3(3) ="抱抱卡"
js3(4) ="补血卡"
js3(5) ="吃豆卡"
js3(6) ="清醒卡"
js3(7) ="补血卡"
js3(8) ="暴豆卡"
js3(9) ="练功卡"
js3(10) ="离婚卡"
js3(11) ="吃豆卡"
js3(12) ="清醒卡"
js3(13) ="抢亲卡"
js3(14) ="真爱卡"
js3(15) ="清纯卡"
js3(16) ="免罪卡"
js3(17) ="财神卡"
js3(18) ="非礼卡"
js3(19) ="福神卡"
js3(20) ="涨钱卡"
js3(21) ="补血卡"
js3(22) ="暴豆卡"
js3(23) ="练功卡"
js3(24) ="福神卡"
js3(25) ="吃豆卡"
js3(26) ="清醒卡"
js3(27) ="补血卡"
js3(28) ="暴豆卡"
js3(29) ="练功卡"
js3(30) ="捣乱卡"
js3(31) ="吃豆卡"
js3(32) ="清醒卡"
js3(33) ="抢亲卡"
js3(34) ="免税卡"
js3(35) ="清纯卡"
js3(36) ="免罪卡"
js3(37) ="财神卡"
js3(38) ="养猪卡"
js3(39) ="福神卡"
js3(40) ="加点卡"
randomize()
myxb = Int(Rnd*41)
zstemp3=add(rs("w5"),js3(myxb),1)
dim js4(28)
js4(0) ="蛊虫"
js4(1) ="归命丸"
js4(2) ="大力丸"
js4(3) ="归命丸"
js4(4) ="大力丸"
js4(5) ="九花玉露丸"
js4(6) ="三花聚顶"
js4(7) ="爆胎易筋丸"
js4(8) ="清凉油"
js4(9) ="大宝日霜"
js4(10) ="后悔药"
js4(11) ="炸药"
js4(12) ="舍利子"
js4(13) ="化尸水"
js4(14) ="蛊虫"
js4(15) ="归命丸"
js4(16) ="大力丸"
js4(17) ="归命丸"
js4(18) ="吸血虫"
js4(19) ="九花玉露丸"
js4(20) ="三花聚顶"
js4(21) ="爆胎易筋丸"
js4(22) ="清凉油"
js4(23) ="大宝日霜"
js4(24) ="后悔药"
js4(25) ="炸药"
js4(26) ="舍利子"
js4(27) ="化尸水"
randomize()
myxc = Int(Rnd*28)
zstemp4=add(rs("w8"),js4(myxc),1)
rs.close
rs.open "select * from 怪物 where 姓名='"&gwm&"'",conn
gwdj=rs("等级")
zt=rs("状态")
gwtl=rs("体力")
lsz=rs("猎杀者")
gjto=rs("攻击")
fyto=rs("防御")
ssto=rs("杀气")
allssto=int(gjto+fyto+ssto)
if dj<gwdj then 
	Response.Write "<script Language=Javascript>alert('提示：你的等级小于怪物等级，无法猎杀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if zt="死亡" then 
	Response.Write "<script Language=Javascript>alert('提示："&gwm&"已被猎杀者"&lsz&"猎杀，想打怪物必须先将其唤醒！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
if allss<allssto then
       ss=0
       ss1=int(allssto-allss)
else
       ss=int(allss-allssto)
       ss1=ss/4
end if
conn.execute "update 怪物 set 体力=体力-"&ss&",杀气=杀气-"&s6&" where 姓名='"&gwm&"'"
conn.execute "update 用户 set 体力=体力-"&ss1&" where 姓名='"&aqjh_name&"'"
rs.open "select * from 怪物 where 姓名='"&gwm&"'",conn
zt=rs("状态")
gwtl=rs("体力")
lsz=rs("猎杀者")
if zt="死亡" then 
	Response.Write "<script Language=Javascript>alert('提示："&gwm&"已被猎杀者"&lsz&"猎杀，想打怪物必须先将其唤醒！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if tl<0 then
conn.execute "update 用户 set 状态='死',事件原因='怪物|壮烈"' where 姓名='" & aqjh_name &"'"
end if
if gwtl<0 then
conn.execute "update 怪物 set 状态='死亡',体力=0,攻击=0,杀气=0,防御=0,猎杀者='"&aqjh_name&"' where 姓名='"&gwm&"'"
conn.execute "update 用户 set 金币=金币+50,allvalue=allvalue+"&s3&",w1='"&zstemp1&"',w6='"&zstemp&"',w3='"&zstemp2&"',w5='"&zstemp3&"',w8='"&zstemp4&"',操作时间=now() where 姓名='"&aqjh_name&"'"
end if
if gwtl<0 then
	Response.Write "<script Language=Javascript>alert('提示："&aqjh_name&"恭喜你！"&gwm&"已经被你成功消灭，获得"&js2(myxa)&"1个，"&js(myxy)&"7个，"&js1(myxx)&"15个，卡片"&js3(myxb)&"1张，配药"&js4(myxc)&"1个，金币50个。积分（存点）"&s3&"点！！');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('经过搏斗，"&gwm&"体力下降"&ss&"点，杀气下降"&s6&"点，你的体力下降"&ss1&"点！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=RED><b>〖梦幻魔界消息〗</b></font><font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]经过一场苦战，已将[<font color=red><b>"&gwm&"</b></font>]消灭，并且从怪物身上获得[<font color=red><b>"&js(myxy)&"</b></font>]7个、[<font color=red><b>"&js1(myxx)&"</b></font>]15个、[<font color=red><b>"&js2(myxa)&"</b></font>]1个、[<font color=red><b>"&js3(myxb)&"</b></font>]1张、配药[<font color=red><b>"&js4(myxc)&"</b></font>]1个和存点[<font color=red><b>"&s3&"</b></font>]点，金币[<font color=red><b>50</b></font>]个！<font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]的<font color=red><b>【会员等级】</b></font>为[<font color=red><b>"&hydj&"</b></font>]级！他进入梦幻魔界的<font color=red><b>【时间限制】</b></font>为：[<font color=red><b>"&jgsj&"</b></font>]秒</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
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
