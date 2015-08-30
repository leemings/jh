<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/sjfunc.asp"-->
<%
pyyl=10  '点给朋友需要的银币数
pyyls=1000000  '点给朋友需要的银两数
xyyl=3 '点给大家需要金币数
xyyls=10000000 '点给大家需要银两数
xzsj=true   '给全体点歌时候的时间限制,如果不开启请该成[false]
xzsjs=5      '给全体点歌时候的分钟数限制
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" or Session("sjjh_inthechat")<>"1" then Response.Redirect "chaterr.asp?id=001"
song=Request.Form("song")
if song="" or song=null then
Response.Write "<script language=JavaScript>{alert('提示：请选择你要点播的歌曲');}</script>"
Response.End
end if
loopok=Request.Form("loopok")
to1=trim(request.form("to1"))
zhufu=trim(request.form("zhufu"))
if zhufu<>"" and zhufu<>null then zhufu=bdsays(zhufu)
Response.write "<html><head></head><body oncontextmenu=self.event.returnValue=false>"
songname=song
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("eline__SONG.mdb")
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
Session("sjjh_lasttime")=sj
Response.Write "<bgsound src=" & chr(34) & song & chr(34) & " loop=" & chr(34) & loopok & chr(34) & "><script Language=JavaScript>parent.f2.startnosay();</script>"
case "ddj"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 金币,银两 from 用户 where 姓名='" & sjjh_name &"'",conn
if xzsj then 
sj=DateDiff("n",Application("sjjh_zxdg"),now())
if sj<xzsjs and sjjh_name<>"回首当年" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你先等一会吧~~~有人点歌了');}</script>"
	Response.End
end if
end if
if rs("金币")<xyyl and sjjh_name<>"回首当年" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要金币["&xyyl&"]才能点歌');}</script>"
	Response.End
end if
if rs("银两")<xyyls and sjjh_name<>"回首当年" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要银两["&xyyls&"]才能点歌');}</script>"
	Response.End
end if
conn.execute "update 用户 set 金币=金币-"&xyyl&" where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&xyyls&" where 姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Application("sjjh_zxdg")=now()
says="<bgsound src=" & chr(34) & song & chr(34) & " loop=1><font color=green>【点歌】</font><font color=blue>"&sjjh_name&"</font><B><font color=#ff0000>点了一首<font color=blue>["&songname&"]</font>给大家，并说:<font color=blue>["&zhufu&"]</font></font></b>"
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
conn.open Application("sjjh_usermdb")
rs.open "select 银币,银两 from 用户 where 姓名='" & sjjh_name &"'",conn
if rs("银币")<pyyl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你需要银币["&pyyl&"]才能点歌');}</script>"
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

conn.execute "update 用户 set 银币=银币-"&pyyl&" where 姓名='"&sjjh_name&"'"
conn.execute "update 用户 set 银两=银两-"&pyyls&" where 姓名='"&sjjh_name&"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
says="<bgsound src=" & chr(34) & song & chr(34) & " loop=1><font color=green>【点歌】</font><font color=blue>"&sjjh_name&"</font><B><font color=#ff0000>点了一首<font color=blue>["&songname&"]</font>送给您，并对你说<font color=blue>["&zhufu&"]</font></font></font></b>"
act="消息"
towhoway=1
towho=to1
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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