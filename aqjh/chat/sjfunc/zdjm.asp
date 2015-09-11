<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.ASP"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_jhdj<18 then
	Response.Write "<script language=JavaScript>{alert('提示：领取存点需要大于18级才可以操作！');}</script>"
	Response.End
end if
temp=int(abs(clng(Application("aqjh_zdff"))))
if temp<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	response.end
end if
zzfff=Application("aqjh_zzfafang")
if zzfff="" then
	Response.Write "<script language=JavaScript>{alert('【江湖】提示：目前没有站长发放。');}</script>"
	Response.End 
end if
fffdata=split(zzfff,"|")
ffsl=abs(int(clng(fffdata(0))))
ffsj=fffdata(1)
erase fffdata
nowsj=DateDiff("s",ffsj,now())
if nowsj>=20 then
        Application.Lock
	Application("aqjh_zzfafang")=""
        Application.UnLock
	Response.Write "<script language=JavaScript>{alert('提示：站长发放有效时间20秒，过期作废。');}</script>"
	Response.End
end if
name=aqjh_name
myip=Request.ServerVariables("REMOTE_ADDR")
if application("aqjh_zzffjl")<>"" then
	ff=split(application("aqjh_zzffjl"),";")
	jsq=ubound(ff)-1
	for i=0 to jsq
		fsz=split(ff(i),"|")
		if fsz(0)=sjjh_name and clng(fsz(1))=hour(time()) then
			erase fsz
			erase ff
			Response.Write "<script language=JavaScript>{alert('提示：你已经领过存点了。');}</script>"
			Response.End
		end if
		if clng(fsz(1))=hour(time()) and fsz(2)=myip then
			erase fsz
			erase ff
			Response.Write "<script language=JavaScript>{alert('提示：对不起，同一IP只能领取一次金币。');}</script>"
			Response.End
		end if
		erase fsz
	next
	erase ff
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,allvalue from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
xianzaisj=s & ":" & f & ":" & m
tt="<font class=t>(" & xianzaisj & ")</font>"
woczsj=DateDiff("s",rs("操作时间"),now())
if  woczsj<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&aqjh_name &"]你刚才领过发放了，不能再要了！');window.close();</script>"
	response.end
end if
newjl=application("aqjh_zzffjl")&aqjh_name&"|"&hour(time())&"|"&myip&";"
application.lock()
application("aqjh_zzffjl")=newjl
application.unlock()
fffff="<marquee width=100% scrollamount=8><font color=green>【爱情发放】</font><font color=red>"&aqjh_name&"</font>领到了站长发放的<font color=red><b>50点存点</b>，站长支持在线，真正玩游戏，不要开小差呀，等会还会发！</font>！</marquee>"&tt
conn.execute "update 用户 set allvalue=allvalue+50,操作时间=now()+(31/86400) where 姓名='"&aqjh_name&"'"
conn.close
set conn=nothing
says=fffff
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
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
%>