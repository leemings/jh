<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#INCLUDE FILE="midsound.asp" -->
<%Response.Expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
towho=Trim(Request.Form("towho"))
if Instr(useronlinename," "&towho&" ")=0 then Response.Redirect "../error.asp?id=410&name=" & server.URLEncode(towho)
song=Request.Form("song")
infofo=Left(Trim(Request.Form("infofo")),100)
infofo=Server.HTMLEncode(Replace(infofo,"%%",mid(song)))
linkurl="mid/"& song &".mid"
linkurl2="mid/" & Server.URLEncode(song) &".mid"
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
Session("sjjh_lasttime")=sj
wbq=Application("sjjh_webicq")
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
wbq(wbqub+3)=sjjh_name
'wbq(wbqub+4)="<div align=center><a href=" & linkurl & ">[IE4]收听</a>　<a href=" & linkurl2 & ">[IE5]收听</a></div>" & infofo
wbq(wbqub+4)="<div align=center><bgsound src="& linkurl &"></div>" & infofo
wbqub=wbqub+4
webicqname=""
for i=1 to wbqub step 4
	webicqname=webicqname & " " & wbq(i+1)
next
webicqname=webicqname&" "
Application.Lock
Application("sjjh_webicq")=wbq
Application("sjjh_webicqname")=webicqname
Application.UnLock
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%>
<html>
<head>
<title>好歌赠好友</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<div align="center"><font color="#FFFF00" style="font-size:12pt">发送完毕</font></div>
<hr size=1 color=FFFF00><br>
<p>已经将：</p>
<p><font color=FFFF00><%=mid(song)%></font></p>
<p>发送给：</p>
<p><font color=FFFF00><%=towho%></font></p>
<p>祝福话语：</p>
<p><font color=FFFF00><%=infofo%></font></p>
<input type="button" name="abort" value="返回" onClick="javascript:history.go(-2)">
</body>
</html>