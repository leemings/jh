<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'小孩
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【小孩出生】<font color=" & sayscolor & ">"+xiaohai(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xiaohai(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select boy from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if  rs("boy")<>"" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：计划生育,只生一个呀！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select 性别,配偶 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if rs("性别")="男" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你是男的也想生呀！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if isnull(rs("配偶")) or rs("配偶")="无" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：对象都没有，想自己一个人生呀！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select 银两,金币,体力,道德 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if rs("银两")<1000000 and rs("金币")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你银两不足100万或者金币没有10个！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("体力")<10000 and rs("道德")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你体力不足1万或者道德不足1千！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select 结婚记念日,配偶 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
if DateDiff("d",rs("结婚记念日"),date())<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的爱情还不够稳定，等2天后吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zz=rs("配偶")
huaname=trim(fn1)
yin=1000000
jinbi=10
tili=10000
dd=1000
randomize
tt=(int(rnd()*999)+51)
if tt mod 2=0 then
   huasex="男"
   boysex="images/boy.gif"
else
   huasex="女"
   boysex="images/girl.gif"
end if
zstemp=huaname&"|"&huasex&"|"&now()&"|1000|1000|1000|0"&"|"&now()
'金龍|男|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
conn.execute "update 用户 set 银两=银两-" & yin & ",金币=金币-"&jinbi&",体力=体力-"&tili&",操作时间=now(),boy='"&zstemp&"',boysex='"&boysex&"' where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 操作时间=now(),boy='"&zstemp&"',boysex='"&boysex&"' where 姓名='"&zz&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& zz &"','操作','生了个乖宝宝')"
xiaohai="##觉得婚姻太寂寞，经过一番思想斗争，##与<font color=red>"&zz&"</font>亲亲爱爱，经[爱情]计生所所长的批准，生下了一个<font color=red>"&huasex&"婴</font>，名字叫<font color=red>"&huaname&"</font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>