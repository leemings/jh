<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc/func.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
f=Minute(time()) 
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
at1=request.form("at1")
at2=request.form("at2")
at3=request.form("at3")
to1=request.form("to1")
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('您正在“暂离”状态中，请使用“我回来啦”功能解除！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"正在“暂离”状态中，请不要攻击！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('所攻击的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('你不能对NPC使用此操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=aqjh_name
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
if rs("门派")="冰心训练营" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：冰心训练营的人不可以操作！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('官府不可以使用连续技');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('警告：请等"& s &"秒再发招,可别累着！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('坏事作尽，杀人数满，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("等级")<=21 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('连续技需要战斗等级21级以上才可以操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("保护")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('请取消练功保护再操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from 用户 where 姓名='" & to1 & "'",conn,2,2
zstt=rs("转生")
if zstt>=5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方已转生5次了，这招伤不了对方！');}</script>"
	Response.End
end if
if rs("门派")="冰心训练营" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是冰心训练营的人不可以操作！');}</script>"
	Response.End
end if
jhhy=rs("会员等级")
if rs("门派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('不可以对官府人操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("保护")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[to1]再在练功保护中，不能操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("等级")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[to1]为初少江湖新手，不用一这么重的招式吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select 门派 from 用户 where 姓名='" & aqjh_name &"'",conn
mp=rs("门派")
rs.close
rs.open "SELECT * FROM y WHERE a='" & at1 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你并没有["&at1&"]这样的武功呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei1=abs(rs("d"))
wg1=abs(rs("c"))
rs.close
rs.open "SELECT * FROM y WHERE a='" & at2 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你并没有["&at2&"]这样的武功呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei2=abs(rs("d"))
wg2=abs(rs("c"))
rs.close
rs.open "SELECT * FROM y WHERE a='" & at3 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你并没有["&at3&"]这样的武功呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei3=abs(rs("d"))
wg3=abs(rs("c"))
nei=nei1+nei2+nei3
wg=wg1+wg2+wg3
rs.close
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
mygj=rs("攻击")
if rs("内力")<nei or rs("武功")<wg then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('内力不够，发不了连续技！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from 用户 where 姓名='" & to1 & "'",conn
tofy=rs("防御")
killer=(mygj-tofy)+nei+wg
'killer=int(((lxjwg1+lxjgj1)-(lxjwg2+lxjgj2)+nei/10)/5)
'杀伤小于100
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update 用户 set 内力=内力-" & nei & ",武功=武功-"&wg&",操作时间=now() where 姓名='" & aqjh_name &"'"
conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 & "'"
e=""
if rs("体力")<-100 then
	conn.execute "update 用户 set 杀人数=杀人数+1 where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set 状态='死,事件原因='"&aqjh_name&"|连续技:"&at1&at2&at3&"'' where 姓名='" & to1 & "'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','连续技','人命')"
	e="点，" & to1 & "慢慢的倒了下去……从此江湖上又少了一只大虾"
	stunt=aqjh_name & "运足内力对" & to1 & "使用了【" & at1 & "+" & at2 & "+" & at3 & "】的一连串华丽连续技，杀伤" & killer & e
	call boot(to1,"连续技，操作者："&aqjh_name&","&at1&at2&at3)
else
	stunt=aqjh_name & "运足内力对" & to1 & "使用了【" & at1 & "+" & at2 & "+" & at3 & "】的一连串华丽连续技，杀伤" & killer & e
end if
says="<font color=red>【连续技】"& stunt &"</font>"& t			'聊天数据

says=replace(says,chr(39),"\'")
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
Response.Write "<script Language=Javascript>alert('恭喜，您的连续技已经发招完成！');location.href = 'javascript:history.go(-1)';</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
%>
