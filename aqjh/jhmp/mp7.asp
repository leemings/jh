<%@ LANGUAGE=VBScript codepage ="936" %><script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
<%Response.Buffer=true

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=Trim(Request.QueryString("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间,等级 from 用户 where 姓名='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<30 then
	s=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "想要给徒弟发武功/内力请等["&s&"]分钟现来！师傅指导：一定要战斗等级30级以上！"
	Response.End
end if
conn.Execute "update 用户 set 操作时间=now() where 姓名='" & aqjh_name &"'"
dj=rs("等级")
rs.close
smoney=0
if dj>=30 then
	smoney=1
end if
if smoney=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "不好意思，你的等级不够不能发武功/内力！师傅指导：一定要战斗等级30级以上！"
	Response.End
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)

online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
randomize timer
	s=int(rnd*20000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>【师傅指导】</font>"& aqjh_name &"<B>亲自给聊天室本门派弟子指导武功，聊天中的弟子武功/内力增加：<font color=#ff0000>"& s &"点！</font></b>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
for i=0 to x
	conn.Execute "update 用户 set 武功=武功+" & s & ",内力=内力+" & s & " where 姓名='" & online(i) & "' and 师傅='"& aqjh_name &"'"
next
conn.close
set conn=nothing
says=diaox

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
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
Response.Redirect "../ok.asp?id=705"
%>