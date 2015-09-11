<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'小孩护主
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=Application("aqjh_onlinekill")
if onlinenow<onlinekill and chatinfo(0)<>"恩怨情仇" then
	Response.Write "<script language=JavaScript>{alert('提示:大厅在线低于"&onlinekill&"人不得动武！');}</script>"
	Response.End
end if
next
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
if f>Application("aqjh_pktime") then
        Response.Write "<script language=JavaScript>{alert('提示：江湖PK打架时间为每小时前["&Application("aqjh_pktime")&"]分，打架请去恩怨！');}</script>"
	response.end
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
says=cwgj(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'小孩护主
function cwgj(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="通吃江湖"  then
Response.Write "<script Language=Javascript>alert('提示：[通吃江湖]房间里现在是只给堂主和护法、长老、掌门等进行门派大战的，其他人等在场可让你门派加强，想打架到[高手决战]房间去吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
fn1=trim(fn1)
rs.open "select boysex,boy,配偶 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
mypic="<img src=/chat/boy/"&rs("boysex")&" width=30 height=30>"
cwnl="武功"
cwgj="攻击"
cwfy="防御"
zz=rs("配偶")
rs.close
rs.open "select 门派,等级,保护,grade,宝物,死亡时间 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("门派")="基督徒" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是教会中人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚被别人杀死，先放过他吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
if rs("等级")<=18 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对初入江湖新手操作！');}</script>"
	Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,grade,保护,杀人数,boy from 用户 where  姓名='"&aqjh_name&"'",conn,2,2
if rs("门派")="基督徒" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是教会中人不可以操作！');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要发招!');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再发招了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("boy")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，赶快生一个！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if fn1<>zt(0) then
	Response.Write "<script language=JavaScript>{alert('提示：你怎么叫自己的孩子都叫错！');}</script>"
	Response.End
end if
cssj=clng(DateDiff("d",zt(6),now()))
gs=clng(zt(6))
if gs<(80+cssj*5) then
	czzt="叛逆"
elseif gs>((100+cssj*5+50)) then
	czzt="顺从"
else
	czzt="普通"
end if
if czzt="叛逆" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩心情不好，不想与你一起打架！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

if clng(zt(3))<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩体力虚弱，不能打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(4))<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩内力不够，不能打架了！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
cwsr=clng(DateDiff("d",zt(2),now()))
if cwsr<10 then
	cwgj="<font color=green>【小孩护主】<font color=" & saycolor & ">##想使用小孩护主攻击%%，你的宝贝还在哺乳期，你忍心吗？(说明：小孩护主攻击计算,小孩出生天数除以5取整数，再去除攻击特性来攻击对方，攻击防御武功值给夫妻俩平均增加!)</font>"
	exit function
end if
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"-"&clng(zt(3))*2&","&cwgj&"="&cwgj&"-"&clng(zt(4))*2&","&cwfy&"="&cwfy&"-"&clng(zt(5))*2&" where 姓名='" & to1 &"'")
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&clng(zt(3))&","&cwgj&"="&cwgj&"+"&clng(zt(4))&","&cwfy&"="&cwfy&"+"&clng(zt(5))&",boy='' where 姓名='" & aqjh_name &"'")
conn.Execute ("update 用户 set "&cwnl&"="&cwnl&"+"&clng(zt(3))&","&cwgj&"="&cwgj&"+"&clng(zt(4))&","&cwfy&"="&cwfy&"+"&clng(zt(5))&",boy='' where 姓名='" & zz &"'")
conn.Execute ("update 用户 set 体力=体力-"&clng(zt(3))*5&" where 姓名='" & to1 &"'")
rs.open "select 体力 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("体力")<-100 then
	call boot(to1,"发招，操作者："&aqjh_name&",["&zt(0)&"]打死"&fn1)
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','使用"&fn1&"同归于尽','人命')"
	conn.execute "update 用户 set 状态='死',登录=now(),事件原因='" & aqjh_name & "|使用"&fn1&"同归于尽' where 姓名='"&to1&"'"
	cwgj="<font color=green>【小孩同归于尽】<font color=" & saycolor & ">##的小孩<b>"&zt(0)&"</b>]"&mypic&"实在看不惯%%，愤起攻击,使%%体力下降:<font color=red>"&clng(zt(3))*5&"</font>点,攻击下降:<font color=red>"&clng(zt(4))*2&"</font>点,防御下降:<font color=red>"&clng(zt(5))*2&"</font>点,小孩特技攻击,%%被吸走[<b>"&cwnl&"</b>]:-"&int(clng(zt(3))*2)&"点,%%终于体力不支倒地，被["&zt(0)&"]打得七零八落………小孩也因此舍身取义……</font>"
else
	cwgj="<font color=green>【小孩同归于尽】<font color=" & saycolor & ">##的小孩<b>"&zt(0)&"</b>]"&mypic&"实在看不惯%%，愤起攻击,使%%体力下降:<font color=red>"&clng(zt(3))*5&"</font>点,攻击下降:<font color=red>"&clng(zt(4))*2&"</font>点,防御下降:<font color=red>"&clng(zt(5))*2&"</font>点,小孩特技攻击,%%被吸走[<b>"&cwnl&"</b>]:-"&int(clng(zt(3))*2)&"点,小孩也因此舍身取义……</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>