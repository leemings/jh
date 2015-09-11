<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
inroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 状态,体力 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
if rs("状态")<>"正常" or rs("体力")<-100 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你已经被打死或状态不正常请重新登陆！');location.href = '../exit.asp'}</script>"
	Response.End
end if
conn.Execute "update 用户 set 宝物='无' where  姓名='" & aqjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
if Session("aqjh_inthechat")<>"1" then Response.Redirect "close.asp"
Session("aqjh_inthechat")="0"
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
onlinelist=Application("aqjh_onlinelist"&inroom)
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),aqjh_name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&inroom)=onlinelist
Application("aqjh_useronlinename"&inroom)=Replace(Application("aqjh_useronlinename"&inroom)," " & aqjh_name & " ","")
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
says="<font color=red>【离线了】</font><font color=F08000>" & Replace(Application("aqjh_userout"),"%%","<font color=black>" & aqjh_name & "</font>") & "</font><font class=t>(" & t & ")</font>"

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");</script>"
addmsg saystr
mycd=DateDiff("n",Session("aqjh_savetime"),now())
if mycd>0 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("aqjh_usermdb")
	Session("aqjh_savetime")=now()
	addvalue=mycd
	conn.execute  "update 用户 set allvalue=allvalue+"&addvalue&",mvalue=mvalue+"&addvalue&",lasttime='"&sj&"',lastip='"&ip&"' where 姓名='" & aqjh_name &"'"
	conn.close
	set conn=nothing
end if
Response.Redirect "close.asp"%>
