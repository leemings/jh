<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
towho=Trim(Request.Form("towho"))
if Instr(useronlinename," "&towho&" ")=0 then Response.Redirect "../error.asp?id=400&name=" & server.URLEncode(towho)
intro=Trim(Request.Form("intro"))
if intro="" then Response.Redirect "../error.asp?id=403"
intro=server.HTMLEncode(intro)
intro=lcase(intro)
intro=Replace(intro,"con","f1uck")
intro=Replace(intro,"or","ｏｒ")
intro=Replace(intro,"java","f1uck")
intro=Replace(intro,"www","f2uck")
intro=Replace(intro,"com","f2uck")
intro=Replace(intro,"net","f2uck")
intro=Replace(intro,"61.","f2uck")
intro=Replace(intro,"202.","f2uck")
intro=Replace(intro,"200.","f2uck")
intro=Replace(intro,"Ｗ","f2uck")
intro=Replace(intro,"ｗ","f2uck")
intro=Replace(intro,"jh","f2uck")
intro=Replace(LCase(intro),LCase("http"),"f2uck")
maren=instr(intro,"f2uck")
if maren<>0 then
%>
<script language="vbscript">
alert( "<%=Application("aqjh_chatroomname")%>是不允许打广告的！")
window.close()
</script>
<%
Response.end
end if
'if instr(intro,"www")<>0 or instr(intro,"net")<>0 or instr(intro,"com")<>0 or instr(intro,"xajh")<>0 or instr(intro,"jh")<>0 or or instr(intro,"net")<>0 or instr(intro,"net")<>0
intro=Replace(intro,chr(13)&chr(10),"<br>")
if len(intro)>1024 then Response.Redirect "../error.asp?id=401"
if Instr(intro,"<br><br><br>")<>0 then Response.Redirect "../error.asp?id=402"
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
Session("aqjh_lasttime")=sj
wbq=Application("aqjh_webicq")
wbqub=UBound(wbq)
if wbqub>0 then
Dim wbqnew()
j=1
for i=1 to wbqub step 4
if DateDiff("n",wbq(i),sj)<=10 then
Redim Preserve wbqnew(j),wbqnew(j+1),wbqnew(j+2),wbqnew(j+3)
wbqnew(j)=wbq(i)
wbqnew(j+1)=wbq(i+1)
wbqnew(j+2)=wbq(i+2)
wbqnew(j+3)=wbq(i+3)
j=j+4
end if
next
if j>=4 then
wbq=wbqnew
else
Dim wbqnull(0)
wbq=wbqnull
end if
wbqub=UBound(wbq)
end if
Redim Preserve wbq(wbqub+1),wbq(wbqub+2),wbq(wbqub+3),wbq(wbqub+4)
wbq(wbqub+1)=sj
wbq(wbqub+2)=towho
wbq(wbqub+3)=aqjh_name
wbq(wbqub+4)=intro
wbqub=wbqub+4
webicqname=""
for i=1 to wbqub step 4
webicqname=webicqname & " " & wbq(i+1)
next
webicqname=webicqname&" "
Application.Lock
Application("aqjh_webicq")=wbq
Application("aqjh_webicqname")=webicqname
Application.UnLock
says="<font color=green>【快乐信使】["&aqjh_name&"]---->["&towho&"]</font>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
Response.Redirect "../ok.asp?id=200&name=" & server.URLEncode(towho)%>