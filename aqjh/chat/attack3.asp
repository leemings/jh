<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
chatroomsn=session("nowinroom")	'当前所在房间序号
chatroomname=Application("aqjh_chatroomname"&mychatroomsn)
mm=Application("aqjh_chatroomname"&chatroomsn)
chatroominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(chatroominfo)-1



aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"

nowinroom=session("nowinroom")
online=split((Application("aqjh_useronlinename"&nowinroom))," ")
aqjh_roominfo=split(Application("aqjh_room"),";")
onlinenum=ubound(online)+1

chatroomnum=ubound(aqjh_roominfo)-1

id=trim(Request.QueryString ("id"))
mode=trim(Request.QueryString ("mode"))
for i=0 to chatroomnum	
	ydl=1
	useronlinename=Application("aqjh_useronlinename"&nowinroom)
if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(id)&" ")<>0 then 
Response.Write "<script language=javascript>alert('出错啦！神兽己经复活了！');</script>"
Response.End 
end if
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(id)&" ")=0 then ydl=0

next


if aqjh_name=""  then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有召唤神兽权利！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select * from 用户 where 姓名='"&aqjh_name&"'"),conn
myzs=rs("转生")
zaohuan1=rs("召唤兽1")
if zaohuan1=""   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有神兽,不要捣乱！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
if zaohuan1="无"   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有神兽,不要捣乱！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
if rs("银两")<10000000   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有1000万现金！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
if rs("金币")<10   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有10个金币,不要捣乱！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
if rs("法力")<10000   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有10000法力,不要捣乱！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
if aqjh_name=""  or  myzs<2 then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有神兽权利！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
rs.Close
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
sj2=n & "-" & y & "-" & r & " " & sj
rs.Open ("select * from 用户 where 姓名='"&id&"' "),conn
if rs.BOF or rs.EOF then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  你的神兽被非法改名了，或者数据出错~ ！');}</script>"
Response.End 
end if
if rs("状态")<>"正常" then 
xinxi="由于神兽被别人陷害,主人"&aqjh_name&"花费了300法力和5个金币复活了神兽"
conn.Execute ("update 用户 set 法力=法力-300,金币=金币-5 where 姓名='"&aqjh_name&"'")
else
xinxi="神兽在睡梦中被唤醒!~激发的能量就要被释放了"
end if
if rs.BOF or rs.EOF then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  对不起！\n  目前无神兽出现！！！\n  请按 [确定] 返 回！');}</script>"
Response.End 
end if




randomize timer
qq=int(rnd()*65*myzs)+1
nn=int(rnd()*5500000*myzs)+1

	ii=int(rnd()*1000000*myzs)+1
	iii=int(rnd()*200000*myzs)+1000
	aa=int(rnd()*1500000*myzs)+1
conn.Execute ("update 用户 set 状态='正常',保护=false,武功="&ii&"+146534,防御="&iii&"+1000,体力="&aa&"+1000,等级="&qq&"*"&myzs&",内力="&nn&"+1000,操作时间=now(),死亡时间='2005-3-15 20:24:45' where 姓名='"&id&"'")
conn.Execute ("update 用户 set 银两=银两-10000000,法力=法力-1000,金币=金币-5 where 姓名='"&aqjh_name&"'")

Session("aqjh_inthechat")="1"
Session("aqjh_savetime")=now()
Session("aqjh_lasttime")=sj
myzanli=0
tjrf=rs("通缉")
jhmp=rs("门派")
dj=rs("等级")
newuser=rs("times")
aqjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
jhtx=rs("名单头像")
sex=rs("性别")
hymd=rs("好友名单")
mywife=rs("配偶")
tili=rs("体力")
nl=rs("内力")
wg=rs("武功")

'myonline = id & "|" & sex & "|" & jhmp & "| npc|" &jhtx& "|" & aqjh_jhdj& "| 0 |"&myzanli&"|"&"正常"
myonline = id &"|" & sex & "|"&aqjh_name&"|被召唤中|" &jhtx& "|"& dj &" |0|9|"&myzanli&"|"&"神兽"
'myonline = id & "|" & sex & "|"&aqjh_name&"| 神兽|" &jhtx& "| 0 |"&myzanli&"|"&"正常"
Application.Lock
'-----------------------------------------------------------------------------------


'这里添加人物进入聊天室在f3显示名单的算法.

dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'按门派名字汉语拼音排序
'if yjl=0 and len(onuser(2))<len(jhmp) then			'按门派名字长度，长的在前
'if yjl=0 and len(onuser(2))>len(jhmp) then			'按门派名字长度,短的在前
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'按名字长度，长的在前
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'按名字长度,短的在前
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if

Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)& " " & id  & " "
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
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
says=replace(says,chr(39),"\'") 
says=replace(says,chr(34),"\"&chr(34)) 

into="爱情江湖大侠<font color=red><b>id:"& aqjh_name &"</b></font>的神兽<font color=blue><b>"& id &"</b></font>被召唤！！{"&xinxi&"}神兽体力"&tili&"点，内力"&nl&"点，武功"&wg&"等级"&dj&"级....."

'-------------------------------------------------------------------------------------------
'这里加入在聊天室显示信息的代码
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
'says="<font color=red>〖npc加入〗</font><font color=green>爱情江湖总站</font><font color=#0000ff>npc</font><img src="&jhtx&" width=60 height=60><s><font color=#0000ff>"& id &"</font></s>来到江湖,大家赶快杀啊!~有经验得哦了"   '聊天数据
says="<font color=black>【召唤神兽】</font><font color=008800>" & Replace(into,"##","<img src="&jhtx&"><a href=javascript:parent.sw(\'[" & id & "]\'); target=f2>" & id &"</a><font color=red><b>id:"& aqjh_id &"</b></font>") & "</font><bgsound src=readonly/cd.mid loop=1>"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & id & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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


'-------------------------------------------------------





%>

