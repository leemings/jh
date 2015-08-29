<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'下毒♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	call mess ("提示：在["&chatinfo(0)&"]房间不可以下毒！",1)
end if
if Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<21) and chatinfo(0)="高手E线" then
	Response.Write "<script language=JavaScript>{alert('提示：现在是夺宝时间，打斗请使用夺宝武功！');}</script>"
	Response.End 
end if
if sjjh_jhdj<=jhdj_duyao then
	call mess ("提示：下毒需要["&jhdj_duyao&"]级才可以操作！",1)
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【下毒】<font color=" & saycolor & ">"+xiadu(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)
'下毒
function xiadu(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	call mess ("提示：滚吧，你想做什么？想捣乱吗？！",1)
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
	call mess ("提示：操作错误，格式如下：[物品名&数量]",1)
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	call mess ("提示：操作错误，数量请使用数字!",1)
end if
zswupin=zt(0)
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	call mess ("提示：物品数量应大于0小于100！",1)
end if
erase zt
'判断杀人数
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="E线江湖"  then
Response.Write "<script Language=Javascript>alert('提示：[E线江湖]房间里现在是只给堂主和护法、长老、掌门等进行门派大战的，其他人等在场可让你门派加强，想下毒到[高手E线]房间去吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select 门派,保护,等级,宝物,grade,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：他是出家人不可以操作！",1)
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and sjjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示："&to1&"刚被别人杀死，饶了他吧…！",1)
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：请不要对官府进行攻击！！",1)
end if
if rs("等级")<=2 and rs("宝物")="无" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：请不要对初入江湖新手操作！",1)
end if
if rs("保护")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：对方正在练功保护请不要偷袭!",1)
end if
rs.close
rs.open "select 门派,保护,杀人数,grade,w2,死亡时间,宝物 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你是出家人不可以操作！",1)
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你刚刚被别人杀死，还是先练练再说！",1)
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你是官府中人请不要发招攻击！",1)
end if
if rs("杀人数")>=int(Application("sjjh_killman")) and sjjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你作恶太多，超过江湖杀人上限"& Application("sjjh_killman") &"不能再下毒了！",1)
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你正在练功保护请不要下毒!",1)
end if
'删除自己的毒药
duyao=abate(rs("w2"),zswupin,wusl)
conn.execute "update 用户 set w2='"&duyao&"' where  姓名='"&sjjh_name&"'"
rs.close
'取出毒药杀伤力
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！",1)
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'用户
conn.execute "update 用户 set 道德=道德-20,体力=体力-" & int(tl/4) & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 内力=内力-" & nl & ",体力=体力-" & tl & " where 姓名='" & to1 &"'"
rs.close
over=""
rs.open "select 体力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='"&sjjh_name&"|下毒:不小心自己中毒了…',死亡时间=now() where 姓名='" & sjjh_name &"'"
	over="##不小心自己也中了剧毒，啊的一声，倒下了…"
	call boot(sjjh_name,"下毒，不小心自己中毒了！")
end if
rs.close
rs.open "select 姓名,体力,宝物 from 用户 where 姓名='" & to1 &"'",conn,2,2
to1=rs("姓名")
if rs("体力")>-100 then
	xiadu="##<bgsound src=wav/anqi.wav loop=1>向%%下["&zswupin&"]共:"&wusl&"个,使%%的内力:-<font color=red>" & nl & "</font>点,体力:-<font color=red>" & tl & "</font>点!环境污染"& sjjh_name &"体力:-<font color=red>"& int((tl)/4 ) &"</font>点!"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function	
end if
conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & sjjh_name &"'"
if rs("宝物")=Application("sjjh_baowuname") then
	conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("sjjh_baowuname") &"' where 姓名='" & sjjh_name &"'"
	conn.execute "update 用户 set 宝物修练=0,宝物='无',保护=true,内力=100,体力=2000 where 姓名='" & to1 &"'"
	xiadu="##把%%的宝物:"& Application("sjjh_baowuname") &"抢走，因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
	xiadu=sjjh_name & "<bgsound src=wav/daipu.wav loop=1>向" & to1 & "偷偷下了毒药["&zswupin&"]共:"&wusl&"个,使" & to1 & "的内力:-" & nl & "体力:-" & tl & "," & to1 & "咽气前说道“查出凶手,为我报仇……”"&over
	conn.execute "update 用户 set 状态='死',事件原因='"&sjjh_name&"|下毒:"&zswupin&"',死亡时间=now() where 姓名='" & to1 &"'"
'记录
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"下毒，操作者："&sjjh_name&","&zswupin)
end function
%>