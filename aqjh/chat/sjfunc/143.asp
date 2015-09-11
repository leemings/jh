<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'绝情刀
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
says="<font color=green>【绝情刀】<font color=" & sayscolor & ">"+jiuqingdao(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function jiuqingdao(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,w6,w9,法力,等级,操作时间 FROM 用户 WHERE  姓名="&"'" & aqjh_name &"'",conn,3,3
fla=rs("法力")
dj=rs("等级")
w6w=rs("w9")
sj=DateDiff("s",rs("操作时间"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('所有有动武倾向的都在每小时的前25分钟，现在是聊天时间！');window.close();}</script>"
	Response.End 
end if
if fla<8000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得8000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<20 then
Response.Write "<script language=JavaScript>{alert('此功能需要20级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 体力,门派 FROM 用户 WHERE 姓名="&"'" & towho &"'",conn,2,2
money=int(rs("体力")/3)
if rs("门派")="官府"  then
Response.Write "<script language=JavaScript>{alert('他是官府人员啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

wuyn=iswp(w6w,"绝情刀")
if wuyn=0 then
	Response.Write "<script language=JavaScript>{alert('你有绝情刀吗？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你想对出家人做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if

conn.execute "update 用户 set 法力=法力-8000,操作时间=now() where 姓名="&"'" & aqjh_name &"'"
fq=abate(w6w,"绝情刀",1)
conn.execute "update 用户 set  w9='"&fq&"' where 姓名='"&aqjh_name&"'"

r=int(rnd*2)+1
select case r
case 2
conn.execute "update 用户 set 体力=体力-"&money&" where 姓名="&"'" & towho &"'"
jiuqingdao="##从腰间拔出一把绝情刀<bgsound src=wav/fx42.wav loop=1>，<img src='../jhmp/gif/wg15.gif' width='44' height='44'>恶狠狠地对着<font color=red>%%</font>刺了下去,受死吧，%%啊的一声，被刺中心脉,曾经沧海难为水啊，体力失去1/3，共计"&money&"点......" 
rs.close
rs.open "select 体力 from 用户 where 姓名="&"'" & towho &"'",conn,2,2
if rs("体力")<=0 then 
	conn.execute "update 用户 set 状态='死' where 姓名="&"'" & towho &"'"
	jiuqingdao="##从腰间拔出一把绝情刀<bgsound src=wav/ZR2199.wav loop=1>，<img src='../jhmp/gif/wg15.gif' width='44' height='44'>恶狠狠地对着<font color=red>%%</font>刺了下去,受死吧,%%啊的一声，由于体力不支，被当场刺死，自古多情空余恨啊..." 
	call boot(towho,"绝情刀，操作者："&aqjh_name&","&fn1)
end if
case else
jiuqingdao="##从腰间拔出一把绝情刀<bgsound src=wav/fx42.wav loop=1>，<img src='../jhmp/gif/wg15.gif' width='44' height='44'>恶狠狠地对着<font color=red>%%</font>刺了下去,受死吧,怎料是爱到深处心恨谁，没刺到啊......"
end select
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>