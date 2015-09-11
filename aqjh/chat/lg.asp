<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
'#####房间处理#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
if chatinfo(7)<>0 then
	Response.Write "<script language=JavaScript>{alert('在["&chatinfo(0)&"]房间不可以练功保护！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT 通缉,门派,杀人数,宝物,保护,操作时间,会员等级 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
hydj=rs("会员等级")
if rs("保护")=false and aqjh_jhdj>100 and hydj=0 then
	Response.Write "<script language=JavaScript>{alert('你的等级已经超过100了，系统限制保护！');}</script>"
	Response.End
end if
if rs("门派")="天网" and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：天网杀手不可以保护！');}</script>"
	Response.End
end if
if rs("门派")="出家"  and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：出家后没有必要保护！');}</script>"
	Response.End
end if
if rs("通缉")=True  and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('逃犯也想保护？！想什么来的！');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman")) and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('作尽天下坏事,还想保护？！');}</script>"
	Response.End
end if
if rs("宝物")=Application("aqjh_baowuname") and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你现在有江湖宝物["& Application("aqjh_baowuname") &"]不能进行练功保护！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<2 then
	s=2-sj
	if rs("保护")=true then
		Response.Write "<script language=JavaScript>{alert('你现在正在练功中，请等["&s&"分钟]再操作！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	else
		Response.Write "<script language=JavaScript>{alert('你刚刚解除练功，请等["&s&"分钟]再操作！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
end if
if rs("保护")=true then
	conn.Execute "update 用户 set 保护=false,操作时间=now() where 姓名='" & aqjh_name &"'"	
	diaox="学艺有成，<img src=xx/gif/wg13.gif>现在可以大展拳脚狠杀一次了，哈哈。。。"
else
	conn.Execute "update 用户 set 保护=true,操作时间=now() where 姓名='" & aqjh_name &"'"
	diaox="为了逃避仇家追杀，进行了练功保护操作，谁也打不了！"
end if
conn.close
set rs=nothing
set conn=nothing
says="<font color=#ff0000>【练功消息】" & aqjh_name & ""& diaox &"</font>"			'聊天数据
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
%>