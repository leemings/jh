<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'抢劫令♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以投掷！');}</script>"
	Response.End
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
says="<font color=green>【抢劫令】<font color=" & saycolor & ">"+qiangjielin(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)

'抢劫令
function qiangjielin(fn1,to1)
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
conn.open Application("sjjh_usermdb")
rs.open "select 门派,法力,银两,存款,保护,等级,宝物,grade,死亡时间,宝物 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死，还是先练练再说！');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对官府进行抢劫令！！');}</script>"
	Response.End
end if
if rs("等级")<=15 and rs("宝物")="无"  and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
if rs("保护")=True   and sjjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,保护,杀人数,grade,w4,死亡时间 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
if rs("门派")="出家" and sjjh_grade<10 then
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
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("sjjh_killman"))   and sjjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & sjjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('你作恶太多，超过江湖杀人上限"& Application("sjjh_killman") &"不能再投掷了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
fla=rs("法力")
if fla<20000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得20000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

if rs("保护")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你正在练功保护请不要投掷!');}</script>"
	Response.End
end if
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
	Response.Write "<script language=JavaScript>{alert('提示：你的物品["&zswupin&"]在系统数据库中并不存在\n请删除此物品或找管理员！！');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'用户
conn.execute "update 用户 set 银两=银两+200000,法力=法力-20000,道德=道德-200,体力=体力-" & int(nl/4) & " where 姓名='" & sjjh_name &"'"
conn.execute "update 用户 set 存款=存款-200000,内力=内力-" & nl & ",体力=体力-" & tl & " where 姓名='" & to1 &"'"
rs.close
over=""

rs.open "select 姓名,体力,宝物 from 用户 where 姓名='" & to1 &"'",conn,2,2
to1=rs("姓名")
if rs("体力")>-100 then
	qiangjielin="##<bgsound src=wav/anqi.wav loop=1>向%%下["&zswupin&"]共:"&wusl&"个,使%%的内力减:<font color=red>-" & nl & "</font>体力减:<font color=red>-" & tl & "</font>!自己却消耗体力:<font color=red>-"& int(nl/4 ) &"</font>点!%%头晕晕的<img src='READONLY/open33.gif' width='50' height='40'>"&over
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if


qiangjielin="##<bgsound src=wav/Phant012.wav loop=1>,对着<font color=red>%%</font>大吼一声使用["&zswupin&"]共:"&wusl&"个:抢劫……,从%%那抢得存款200000两,##消耗法力20000点"
conn.execute "update 用户 set 状态='死',事件原因='"&sjjh_name&"|魔宝:"&zswupin&"' where 姓名='" & to1 &"'"
'记录
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"魔宝，操作者："&sjjh_name&","&zswupin)
end function
%>
