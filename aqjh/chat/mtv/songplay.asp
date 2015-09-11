<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" or Session("aqjh_inthechat")<>"1" then Response.Redirect "../chaterr.asp?id=001"
song=Request.Form("song")
if song="" or song=null then
Response.Write "<script language=JavaScript>{alert('提示：请选择你要点播的歌曲');}</script>"
Response.End
end if
loopok=Request.Form("loopok")
to1=trim(request.form("to1"))
zhufu=trim(request.form("zhufu"))
zhufu=Replace(zhufu,"Ｗ","")
zhufu=Replace(zhufu,"ｗ","7758530")
zhufu=Replace(zhufu,"asp","7758530")
zhufu=Replace(zhufu,"jh","7758530")
zhufu=Replace(zhufu,"com","7758530")
zhufu=Replace(zhufu,"net","7758530")
zhufu=Replace(zhufu,"w","7758530")
zhufu=Replace(zhufu,"qq","7758530")
zhufu=Replace(LCase(zhufu),LCase("http"),"7758530")
zhufu=Replace(zhufu,"腾讯","7758530")
zhufu=Replace(zhufu,"江湖","7758530")
if zhufu<>"" and zhufu<>null then zhufu=bdsays(zhufu)
Response.write "<html><head></head><body oncontextmenu=self.event.returnValue=false>"
songname=Trim(song)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
Set Rs=conn.Execute("select  * FROM song where name='"&song&"'")
if not rs.bof and not rs.eof then
song=rs("url")
else
song=""
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
if Request.Form("play")="播放" then
select case loopok
case "1","infinite" 
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
Response.Write "<bgsound src=" & chr(34) & song & chr(34) & " loop=" & chr(34) & loopok & chr(34) & "><script Language=JavaScript>parent.f2.startnosay();</script>"
case "ddj"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,银两 from 用户 where 姓名='" & aqjh_name &"'",conn
if xzsj then 
sj=DateDiff("n",Application("aqjh_zxdg"),now())
if sj<xzsjs and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你先等一会吧~~~有人点歌了！');}</script>"
	Response.End
end if
end if
if rs("金币")<xyyl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要金币["&xyyl&"]才能点歌');}</script>"
	Response.End
end if
if rs("银两")<xyyls then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要金币银两["&xyyls&"]才能点歌');}</script>"
	Response.End
end if
conn.execute "update 用户 set 金币=金币-"&xyyl&" where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&xyyls&" where 姓名='"&aqjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Application("aqjh_zxdg")=now()
says="<img src=img/music.gif><font color=blue>"&aqjh_name&"</font>(花费了"&xyyl&"个金币)<font color=#ff0000>点了一首<font color=blue>["&songname&"]</font>给大家，并说:<font color=blue>["&zhufu&"]</font></font><input type=button value=收听 onclick=javascript:window.open("&chr(34)&"mtv/dj.asp?song="&song&chr(34)&","&chr(34)&"mp3"&chr(34)&","&chr(34)&"Status=no,scrollbars=yes,resizable=no,width=250,height=80,top=20,left=20"&chr(34)&") style="&chr(34)&"background-color:#86A231;color:FFFFFF;border: 1 double"&chr(34)&">"
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
case "dhy"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币,银两 from 用户 where 姓名='" & aqjh_name &"'",conn
if rs("金币")<pyyl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要金币["&pyyl&"]才能点歌');}</script>"
	Response.End
end if
if rs("银两")<pyyls then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要银两["&pyyls&"]才能点歌');}</script>"
	Response.End
end if
if to1="就你自己一个人" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：用这个不省钱的');}</script>"
	Response.End
end if
conn.execute "update 用户 set 金币=金币-"&pyyl&" where 姓名='"&aqjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&pyyls&" where 姓名='"&aqjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
says="<img src=img/music.gif><font color=blue>"&aqjh_name&"</font><font color=#ff0000>点了一首<font color=blue>["&songname&"]</font>送给您，并对你说<font color=blue>["&zhufu&"]</font></font></font><input type=button value=收听 onclick=javascript:window.open("&chr(34)&"mtv/dj.asp?song="&song&chr(34)&","&chr(34)&"mp3"&chr(34)&","&chr(34)&"Status=no,scrollbars=yes,resizable=no,width=250,height=80,top=20,left=20"&chr(34)&") style="&chr(34)&"background-color:#86A231;color:FFFFFF;border: 1 double"&chr(34)&">"
act="消息"
towhoway=1
towho=to1
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
end select
end if
Response.Write "</body></html>"%>