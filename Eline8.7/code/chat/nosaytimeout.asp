<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if Session("sjjh_inthechat")="1" then
Session("sjjh_inthechat")="0"
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
ip=Request.ServerVariables("REMOTE_ADDR")
'dim newonlinelist()
'************
Application.Lock
onlinelist=Application("sjjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),sjjh_name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
Application("sjjh_useronlinename"&nowinroom)=Replace(Application("sjjh_useronlinename"&nowinroom)," " & sjjh_name & " ","")
Application.UnLock
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
act="退出"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<bgsound src=wav/diao.wav loop=1><font color=black>【默哀】" & sjjh_name &"</font><font color=F08000>由于潜水超过" & Application("sjjh_maxtimeout") & "分钟，不幸沉入海底……</font><font class=t>(" & t & ")</font>"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");</script>"
addmsg saystr
end if
Response.Redirect "chaterr.asp?id=002"%>