<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'投掷
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<=jhdj_aq then
	Response.Write "<script language=JavaScript>{alert('提示：投掷需要["&jhdj_aq&"]级才可以操作！');}</script>"
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
act=0
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
says="<font color=green>【投掷】<font color=" & saycolor & ">"+touzi(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)

'投掷
function touzi(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('提示：滚吧，你想做什么？想捣乱吗？！');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('提示：操作错误，格式如下：[物品名&数量]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('提示：操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
zswupin=zt(0)
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于100！');}</script>"
	Response.End 
end if
erase zt
'判断杀人数
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,保护,等级,宝物,grade,死亡时间,宝物 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死，还是先练练再说！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对官府进行攻击！！');}</script>"
	Response.End
end if
if rs("等级")<=30 and rs("宝物")="无"  and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
if rs("保护")=True   and aqjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,保护,杀人数,grade,w4,死亡时间,杀人时间 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死，还是先练练再说！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("杀人时间"),now())
if sj<30 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚杀人了，等30分钟后再杀人吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman"))   and aqjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再投掷了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你正在练功保护请不要投掷!');}</script>"
	Response.End
end if
'删除自己的暗器
duyao=abate(rs("w4"),zswupin,wusl)
conn.execute "update 用户 set w4='"&duyao&"' where  姓名='"&aqjh_name&"'"
rs.close
'取出暗器杀伤力
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'用户
conn.execute "update 用户 set 道德=道德-20,体力=体力-" & int(nl/4) & " where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 内力=内力-" & nl & ",体力=体力-" & tl & " where 姓名='" & to1 &"'"
rs.close
over=""
rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='"&aqjh_name&"|暗器:不小心暗器伤了自己…',死亡时间=now() where 姓名='" & aqjh_name &"'"
	over="##学艺不精，暗器伤了自己，啊的一声，倒下了…"
	call boot(aqjh_name,"暗器，不小心伤了自己！")
end if
rs.close
rs.open "select 姓名,体力,宝物,杀人时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
to1=rs("姓名")
if rs("体力")>-100 then
	touzi="##<bgsound src=wav/anqi.wav loop=1>向%%下["&zswupin&"]共:"&wusl&"个,使%%的内力减:<font color=red>-" & nl & "</font>体力减:<font color=red>-" & tl & "</font>!自己却消耗体力:<font color=red>-"& int(nl/4 ) &"</font>点!"&over
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
	touzi="##把%%的宝物:"& Application("aqjh_baowuname") &"抢走，因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
touzi="##<bgsound src=wav/daipu.wav loop=1>向%%投掷了["&zswupin&"]共:"&wusl&"个,使%%的内力::<font color=red>-" & nl & "</font>体力::<font color=red>-" & tl & "</font>," & to1 & "咽气前说道“查出凶手,为我报仇……”"&over
conn.execute "update 用户 set 杀人时间=now(),状态='死',事件原因='"&aqjh_name&"|暗器:"&zswupin&"' where 姓名='" & to1 &"'"
'记录
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','"&zswupin&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"暗器，操作者："&aqjh_name&","&zswupin)
end function
%>