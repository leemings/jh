<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
inroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 状态,体力 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
if rs("状态")<>"正常" or rs("体力")<-100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你已经被打死或状态不正常请重新登陆！');location.href = '../exit.asp'}</script>"
	Response.End
end if
conn.Execute "update 用户 set 宝物='无' where  姓名='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
if Session("sjjh_inthechat")<>"1" then Response.Redirect "close.asp"
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
onlinelist=Application("sjjh_onlinelist"&inroom)
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
Application("sjjh_onlinelist"&inroom)=onlinelist
Application("sjjh_useronlinename"&inroom)=Replace(Application("sjjh_useronlinename"&inroom)," " & sjjh_name & " ","")
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
says="<font color=black>【离线】</font><font color=333333>" & Replace(Application("sjjh_userout"),"%%","<font color=green>" & sjjh_name & "</font>") & "</font><font class=t>(" & t & ")</font>"

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");</script>"
addmsg saystr
mycd=DateDiff("n",Session("sjjh_savetime"),now())
if mycd>0 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	Session("sjjh_savetime")=now()
	addvalue=mycd
	conn.execute  "update 用户 set allvalue=allvalue+"&addvalue&",mvalue=mvalue+"&addvalue&",lasttime='"&sj&"',lastip='"&ip&"' where 姓名='" & sjjh_name &"'"
	conn.close
	set conn=nothing
end if
Response.Write "<script Language=Javascript>alert('亲爱的：直到今天你才答应让我这样叫你，想必有一天我要去你那见见你的父母。你给我的 18 封情书和在江湖的千言万语，我一定会好好珍藏。我有空会写信给你的，要记得回信喔，虽然我们已见面多次了，但我现在却又渴望马上再见到你，你一定要来这里看看我哦，我已爱你无法自拔！欢迎再次光临一线网络！\nSEE YOU！');window.parent.close();</script>"
response.end
Response.Redirect "close.asp"%>