<%@ LANGUAGE=VBScript%>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
jifen=rs("allvalue")
if rs("转生")<3  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想使用幽灵盾，必须转生3次！');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("银两")<50000000  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想使用幽灵盾，必须有5千万现金！');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("金币")<10  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想使用幽灵盾，必须至少有5个金币！');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("法力")<200000  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('提示：想使用幽灵盾，必须有法力200000 ！');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<300 then
	ss=300-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('300秒练一次，别累着，再等"&ss&"秒吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
neili=int(jifen*.25)
tili=int(jifen*.25)
wg=int(jifen*.05)
conn.execute "update 用户 set 攻击=攻击+10000,防御=防御+10000,攻击加=攻击加+100,防御加=防御加+200,银两=银两-50000000,金币=金币-10,法力=法力-200000,操作时间=now()  where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing	
set conn=nothing
says="<font color=#ff0000><b>【幽灵盾】爱情江湖" & aqjh_name & "</b>大侠，使用幽灵盾成功，增加攻击10000和防御10000，攻击防御上限上升100和200点,继续努力吧！当转生人就是好咯！！！）</font>"			'聊天数据
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
Response.Write "<script language=JavaScript>{alert('提示：幽灵盾施展完毕，攻击防御大大提升！');location.href = 'javascript:history.go(-1)';}</script>"
response.end
%>
