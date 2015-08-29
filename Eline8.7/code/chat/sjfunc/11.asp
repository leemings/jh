<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'偷钱♀wWw.51eline.com♀
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
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以偷钱！');}</script>"
	Response.End
end if
if sjjh_jhdj<=jhdj_tq then
	Response.Write "<script language=JavaScript>{alert('提示：偷钱需要["&jhdj_tq&"]级才可以操作！');}</script>"
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
says="<font color=green>【盗钱】<font color=" & saycolor & ">"+touqian(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'偷钱
function touqian(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="E线江湖"  then
Response.Write "<script Language=Javascript>alert('提示：[E线江湖]房间里现在是只给堂主和护法、长老、掌门等进行门派大战的，其他人等在场可让你门派加强，想偷钱到[高手E线]房间去吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select 门派,保护,等级,姓名,会员等级,银两,grade,操作时间 from 用户 where 姓名='" & sjjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<int(rnd*3) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=int(rnd*3)-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：出家人不可以操作!');}</script>"
	Response.End
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要偷钱!');}</script>"
	Response.End
end if
rs.close
rs.open "select 门派,保护,等级,姓名,会员等级,银两,grade from 用户 where 姓名='" & to1 &"'" ,conn,2,2
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：不可以对出家人操作!');}</script>"
	Response.End
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷钱!');}</script>"
	Response.End
end if
if rs("等级")<=15 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('初入江湖他会有钱吗？！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你要偷管理员的钱想死呀！？！');}</script>"
	Response.End
end if
to1=rs("姓名")
jhhy=rs("会员等级")
yin=rs("银两")
rs.close
randomize timer
'会员2级以上偷钱成功小于5%！
if jhhy<>0 then
	s=int(rnd*5)+1
else
	s=int(rnd*15)+1
end if
yin=int((yin/100))*s
r=int(rnd*7)
Select Case r
Case 1
	conn.execute "update 用户 set 银两=银两-"&yin&" where 姓名='" & to1 &"'"
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>发了一招飞龙探云手,偷取%%银两"& s &"%,共计:"& yin &"两!但不小心全部丢失了!"
Case 2
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>发了一招飞龙探云手,怎奈##与官府认识,##连忙赔礼道歉,灰溜溜的走了!"
Case 3
	conn.execute "update 用户 set 体力=体力-1500 where 姓名='" & sjjh_name &"'"
	touqian="##<bgsound src=wav/oh_no.wav loop=1>想偷取%%的银两,不过被大家发现了,体力下将<font color=red>1500</font>点!<img src='img/daren.gif'>"
Case 4
	rs.open "select 银两 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
	if rs("银两")>50000 then
		conn.execute "update 用户 set 银两=银两-50000,操作时间=now() where 姓名='" & sjjh_name &"'"
		touqian="##<bgsound src=wav/qt.wav loop=1>偷%%的钱被发现,只好上下打点花了<font color=red>5万</font>块,逃过此劫,幸运呀!"
	else
		conn.execute "update 用户 set 状态='狱',操作时间=now() where 姓名='" & sjjh_name &"'"
		touqian="##<bgsound src=wav/oh_no.wav loop=1>刚刚把手伸进%%的钱袋,就被官府的人发现了,立即被捉去坐牢了"
		call boot(sjjh_name,"偷钱，操作者："&to1&"，让你再偷！")
	end if
	rs.close
case else
	conn.execute "update 用户 set 银两=银两-"&yin&" where 姓名='" & to1 &"'"
	conn.execute "update 用户 set 银两=银两+"&yin&" where 姓名='" & sjjh_name &"'"
	touqian= "##<bgsound src=wav/xiaotou.wav loop=1>发了一招飞龙探云手,偷取%%银两"& s &"%,共计:"& yin &"两!%%大叫,我要告你哟!"
End Select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
