<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
'#####房间处理#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
if chatinfo(7)<>0 then
	Response.Write "<script language=JavaScript>{alert('在["&chatinfo(0)&"]房间不可以练功保护！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT 通缉,门派,杀人数,宝物,保护,操作时间 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
f=Minute(time())
'if f<30 and rs("保护")=true and chatinfo(0)<>"高手房间" then
'	Response.Write "<script language=JavaScript>{alert('格斗时间为每小时的后30分钟，非格斗时间不能出关！');window.close();}</script>"
'	Response.End 
'end if
'if f<30 and rs("保护")=False and chatinfo(0)<>"高手房间" then
'	Response.Write "<script language=JavaScript>{alert('格斗时间为每小时的后30分钟，非格斗时间无需闭关！');window.close();}</script>"
'	Response.End 
'end if
if rs("门派")="天网" and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：天网杀手不可以保护！');}</script>"
	Response.End
end if
if rs("门派")="出家" and rs("保护")=False then
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
if rs("杀人数")>=int(Application("sjjh_killman"))  and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('作尽天下坏事,还想保护？！');}</script>"
	Response.End
end if
if rs("宝物")=Application("sjjh_baowuname")  and rs("保护")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你现在有江湖宝物["& Application("sjjh_baowuname") &"]不能进行练功保护！');}</script>"
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
		Response.Write "<script language=JavaScript>{alert('你刚刚解除保护，请等["&s&"分钟]再操作！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
end if
if rs("保护")=true then
	conn.Execute "update 用户 set 保护=false,操作时间=now() where 姓名='" & sjjh_name &"'"	
	diaox="【出关格斗】<font color=blue>" & sjjh_name & "</font>学艺有成，现在可以尽情格斗了，哈哈。。。别惹我！"
else
	conn.Execute "update 用户 set 保护=true,操作时间=now() where 姓名='" & sjjh_name &"'"
	diaox="【闭关修炼】<font color=blue>" & sjjh_name & "</font>进行闭关修炼，打开了保护，谁也别想动我！"
end if
conn.close
set rs=nothing
set conn=nothing
says="<font color=#cc0000>"& diaox &"</font>"			'聊天数据
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
