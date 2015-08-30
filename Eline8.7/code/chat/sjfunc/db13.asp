<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'夺宝投掷♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
call roompd("夺宝投掷")
if sjjh_jhdj<=jhdj_aq then
	Response.Write "<script language=JavaScript>{alert('提示：投掷需要["&jhdj_aq&"]级才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
if towho=sjjh_name or towho="大家" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('提示：不可以对自已或对机器人、大家使用！');}</script>"
	Response.End
end if
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【夺宝投掷】<font color=" & saycolor & ">"+touzi(mid(says,i+1),towho)+"</font>"
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
if wusl=0 or wusl>30 then
	Response.Write "<script language=JavaScript>{alert('提示：物品数量应大于0小于30！');}</script>"
	Response.End 
end if
erase zt
'判断杀人数
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,grade from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对官府进行攻击！！');}</script>"
	Response.End
end if
tonameid=rs("id")
rs.close
rs.open "select id,门派,保护,杀人数,grade,w4,死亡时间 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("grade")>=6 and sjjh_name<>application("sjjh_user") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
mynameid=rs("id")
'删除自己的暗器
duyao=abate(rs("w4"),zswupin,wusl)
conn.execute "update 用户 set w4='"&duyao&"' where  姓名='"&sjjh_name&"'"
rs.close
'取出暗器杀伤力
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'用户
conn.execute "update 用户 set 道德=道德-20,体力=体力-" & int(nl/4) & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 内力=内力-" & nl & ",体力=体力-" & tl & " where 姓名='" & to1 &"'"
rs.close
over=""
rs.open "select 体力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',事件原因='"&sjjh_name&"|暗器:不小心暗器伤了自己…',死亡时间=now() where 姓名='" & sjjh_name &"'"
	over="##学艺不精，暗器伤了自己，啊的一声，倒下了…"
	newover=sjjh_name&"学艺不精，暗器伤了自己，啊的一声，倒下了…"
	call boot(sjjh_name,"暗器，不小心伤了自己！")
end if
rs.close
rs.open "select 姓名,体力 from 用户 where 姓名='" & to1 &"'",conn,2,2
to1=rs("姓名")
if rs("体力")>-100 then
	touzi="##<bgsound src=wav/anqi.wav loop=1>向%%下["&zswupin&"]共:"&wusl&"个,使%%的内力减:<font color=red>-" & nl & "</font>体力减:<font color=red>-" & tl & "</font>!自己却消耗体力:<font color=red>-"& int(nl/4 ) &"</font>点!"&over
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if newover<>"" then
		bbb=zxrsbd(toname,tonameid)
		newtouzi=replace(touzi,"##",sjjh_name)
		newxiadu=replace(newtouzi,"%%",toname)
		touzi=touzi & over & "<br><br>"&bbb
		newover=newxiadu&"<br><br>"&bbb
		call dbxx(newover)
		to1=toname
	end if
	exit function
end if
conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & sjjh_name &"'"
touzi="##<bgsound src=wav/daipu.wav loop=1>向%%投掷了["&zswupin&"]共:"&wusl&"个,使%%的内力::<font color=red>-" & nl & "</font>体力::<font color=red>-" & tl & "</font>," & to1 & "咽气前说道“查出凶手,为我报仇……”"&over
conn.execute "update 用户 set 状态='死',事件原因='"&sjjh_name&"|暗器:"&zswupin&"' where 姓名='" & to1 &"'"
'记录
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"暗器，操作者："&sjjh_name&","&zswupin)
bbb=zxrsbd(sjjh_name,mynameid)
newover=touzi&"<BR><BR>"&bbb
touzi=touzi&"<BR><BR>"&bbb
call dbxx(newover)
to1=toname
end function
%>