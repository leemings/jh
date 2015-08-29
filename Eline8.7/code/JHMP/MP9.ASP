<%@ LANGUAGE=VBScript codepage ="936" %><script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
<%Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
id=Trim(Request.QueryString("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 操作时间,等级,身份,门派,grade from 用户 where 姓名='" & sjjh_name &"'",conn
if rs("grade")<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不是掌门不可以发钱！');}</script>"
	Response.End
end if
mp=rs("门派")
sj=DateDiff("n",rs("操作时间"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "想要给本门弟子发武功/内力请等["&s&"]分钟现来！"
	Response.End
end if
conn.Execute "update 用户 set 操作时间=now() where 姓名='" & sjjh_name &"'"
dj=rs("等级")
rs.close
smoney=0
if dj>=40 then
	smoney=1
end if
if smoney=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "你的等级不够不能发武功/内力！指导弟子要战斗等级40以上！"
	Response.End
end if
useronlinename=Application("sjjh_useronlinename"&nowinroom)

online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
randomize timer
	s=int(rnd*20000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>【指导弟子】</font><font color=blue>["&mp&"]</font>掌门"& sjjh_name &"亲自给聊天室本门派弟子指导武功，聊天中的弟子武功/内力增加：<B><font color=#ff0000>"& s &"点！</font></b>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
for i=0 to x
	conn.Execute "update 用户 set 武功=武功+" & s & ",内力=内力+" & s & " where 姓名='" & online(i) & "' and 门派='"&mp&"'"
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
Response.Redirect "../ok.asp?id=705"
%>