<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_roominfo=split(Application("aqjh_room"),";")

chatroomname=trim(Application("aqjh_chatroomname"&session("nowinroom")))
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(aqjh_name)&" ")=0 then ydl=0

next
id=trim(Request.QueryString ("id"))
'mode=trim(Request.QueryString ("mode"))

if aqjh_name<>"回首当年" and aqjh_name<>"憨豆" then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  你没有发放老朋友的权利！！！\n  请按 [确定] 返 回！');}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
'rs.Open ("select s60,s62,s64,s66 from s"),conn
select case mode
case "僵尸王"
p=int(1000000)
color="#FF3333"
case "大僵尸"
p=int(800000)
color="#FFFF66"
case "小僵尸"
p=int(400000)
color="#00FFFF"
case "小可爱僵尸"
p=int(100000)
color="#FF66FF"
end select
'rs.Close
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
Response.Write "<script language=JavaScript>{alert('  对不起！\n  目前无npc出现！！！\n  请按 [确定] 返 回！');}</script>"
Response.End 
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)
if InStr(LCase(useronlinename)," " & LCase(id) & " ")<>0 then 
Response.Write "<script language=javascript>alert('出错啦！此人己经复活了！');</script>"
Response.End 
end if

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

myonline = id &"|" & sex & "|"&jhmp&"|"&jhsf&"|" & jhtx & "|"& dj&" |"&aqjh_id&"|1|"&myzanli&"|"&""
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
if yjl=0 and StrComp(onuser(0),id,1)=1 then	'按名字汉语拼音排序
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


'-------------------------------------------------------------------------------------------
'这里加入在聊天室显示信息的代码
says="<font color=red>〖公告〗</font><font color=green>快乐江湖</font><font color=#0000ff>大侠</font><img src="&jhtx&"><font color=#0000ff>"& id &"</font><font color=red>id</font>:"&aqjh_id&"乘坐着江湖公交的<IMG src=../hcjs/jhjs/images/tr003.gif>来到了快乐江湖，一抬头看到众多的帅哥美女，急忙拱手曰：“众位大虾，小的有礼了！”"   '聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
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

