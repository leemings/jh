<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sayRandom.asp"-->
<%Response.Buffer=true
sjjh_name=Session("sjjh_name")
inroom=session("nowinroom")
Response.Write "<script language=javascript>if(window==window.top){top.location.href='chaterr.asp';}</script>"
if sjjh_name="" then Response.Redirect "chaterr.asp"
useronlinename=Application("sjjh_useronlinename"&inroom)
if Instr(LCase(useronlinename),LCase(" "&Session("sjjh_name")&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
 
if DateDiff("n",Session("sjjh_savetime"),now())>=15 then
	Response.Write "<script Language=Javascript>parent.clsok=1;if(parent.hang>500){parent.qp();}parent.f3.location.href='savevalue2.asp';</script>"
end if
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
t="<font class=t>(" & sj & ")</font>"
says="<font color=red>【泡点精灵】[</font><font color=#000000>" & sjjh_name & ""& "</font><font color=red>]自动泡点中．．．所得积分减半</font><font color=#ff00ff>↓</font><font color=red>提倡人工存点</font><font color=#ff00ff><b>↑</b></font><font color=red>"&t&"</font>"			'聊天数据
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& inroom & ");<"&"/script>"
addmsg saystr

RandomSay

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
%>

