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
rs.open "select 操作时间,门派,grade,身份,配偶 from 用户 where 姓名='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<30 then
	s=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：想要给本门弟了发钱请等["&s&"]分钟现来！所能发钱人[内力/体力]：掌门、掌门夫人[丈夫]、及本门长老！');window.close();</script>"
	Response.End
end if
conn.Execute "update 用户 set 操作时间=now() where 姓名='" & aqjh_name &"'"
fmp=rs("门派")
fgrade=rs("grade")
fsf=rs("身份")
fpo=rs("配偶")
rs.close
smoney=0
if (fgrade>=5 and fsf="掌门") or (fgrade=4 and fsf="长老") then
	smoney=1
end if
if fpo<>"无" then
rs.open "select 门派,grade,身份,配偶 from 用户 where 姓名='"& fpo &"'",conn
	if rs("门派")=fmp and rs("grade")=5 and rs("身份")="掌门" then
		smoney=1
	end if
rs.close
end if
if smoney=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：不好意思，你的身份不能发钱[内力/体力]！所能发钱人[内力/体力]：掌门、掌门夫人[丈夫]、及本门长老！');window.close();</script>"
	Response.End
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
randomize timer
if id=1 then
	s=int(rnd*5000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>【"&fmp&"武功指导】</font>"& aqjh_name &"<B>亲自给聊天室本门派弟子指导武功，聊天中的弟子内力/体力增加：<font color=#ff0000>"& s &"点！</font></b>"
else
	s=int(rnd*200000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>【"&fmp&"发钱】</font>"& aqjh_name &"<B>给自己门派弟子每一个人发放江湖银两：<font color=#ff0000>"& s &"两<img src=img/251.gif>！</font></b>"
end if

for i=0 to x
if id=1 then
	conn.Execute "update 用户 set 体力=体力+" & s & ",内力=内力+" & s & " where 姓名='" & online(i) & "' and 门派='"& fmp &"'"
else
	conn.Execute "update 用户 set 银两=银两+" & s & " where 姓名='" & online(i) & "' and 门派='"& fmp &"'"
end if
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