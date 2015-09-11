<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'偷金币
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
if aqjh_jhdj<=jhdj_tq then
	Response.Write "<script language=JavaScript>{alert('提示：偷金币需要["&jhdj_tq&"]级才可以操作！');}</script>"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【偷金币】<font color=" & saycolor & ">"+touqian(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'偷金币
function touqian(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 保护,等级,姓名,会员等级,金币,grade,职业,操作时间 from 用户 where 姓名='" & aqjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("操作时间"),now())
if sj<20 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=20-sj
	Response.Write "<script language=JavaScript>{alert('提示：请你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要偷金币!');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要偷金币！！！');}</script>"
	Response.End
end if
if rs("职业")<>"小偷" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的职业不是小偷，所以您不能在这里为他人进行偷金币！！');window.close();</script>"
	response.end
end if
rs.close
rs.open "select 保护,等级,姓名,会员等级,金币,grade from 用户 where 姓名='" & to1 &"'" ,conn,2,2
if rs("保护")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷金币!');}</script>"
	Response.End
end if
if rs("等级")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('初入江湖他会有金币吗？！');}</script>"
	Response.End
end if
if rs("金币")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('他已经没几个金币了,你就辛辛好放过他吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你要偷管理员的金币想死呀！？！');}</script>"
	Response.End
end if
to1=rs("姓名")
jhhy=rs("会员等级")
yin=rs("金币")
rs.close
randomize timer
'会员2级以上偷金币成功小于5%！
if jhhy<>0 then
	s=int(rnd*4)+1
else
	s=int(rnd*4)+1
end if
yin=int((yin/50))*s
r=int(rnd*7)
Select Case r
Case 1
	conn.execute "update 用户 set 金币=金币-"&yin&" where 姓名='" & to1 &"'"
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>发了一招<img src=xx/gif/WG17.GIF>飞龙探云手,偷取%%金币"& s &"%,共计:"& yin &"个!但不小心全部丢失了!"
Case 2
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>发了一招飞龙探云手,怎奈##与官府认识,##连忙赔礼道歉,灰溜溜的走了!"
Case 3
	conn.execute "update 用户 set 体力=体力-1500 where 姓名='" & aqjh_name &"'"
	touqian="##<bgsound src=wav/oh_no.wav loop=1>想偷取%%的金币,不过被大家发现了,体力下将<font color=red>1500</font>点!<img src='img/daren.gif'>"
Case 4
	rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
	if rs("体力")>5000 then
		conn.execute "update 用户 set 体力=体力-5000 where 姓名='" & aqjh_name &"'"
		touqian="##<bgsound src=wav/qt.wav loop=1>偷%%的金币被发现,被大家发现揍了一顿,体力下降了<font color=red>5000</font>点,逃过此劫,幸运呀!"
	else
		conn.execute "update 用户 set 状态='狱' where 姓名='" & aqjh_name &"'"
		touqian="##<bgsound src=wav/oh_no.wav loop=1>刚刚把手伸进%%的金币袋,就被官府的人发现了,立即被捉去坐牢了"
		call boot(aqjh_name,"偷金币，操作者："&to1&"，让你再偷！")
	end if
	rs.close
case else
	conn.execute "update 用户 set 金币=金币-"&yin&" where 姓名='" & to1 &"'"
	conn.execute "update 用户 set 金币=金币+"&yin&" where 姓名='" & aqjh_name &"'"
	touqian= "##<bgsound src=wav/xiaotou.wav loop=1>发了一招<img src=xx/gif/WG19.GIF>飞龙探云脚,偷取%%金币"& s &"%,共计:"& yin &"个!%%大叫,我要告你哟!"
End Select
conn.execute "update 用户 set 操作时间=now() where 姓名='" & aqjh_name &"'"
set rs=nothing	
conn.close
set conn=nothing
end function
%>