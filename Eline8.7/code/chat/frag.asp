<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fromname=LCase(trim(request.querystring("fromname")))
toname=LCase(trim(request.querystring("toname")))
if toname<>sjjh_name then
	if fromname<>toname then
		Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"也没有要替你清理粪库说！');}</script>"
		Response.End
	end if
		Response.Write "<script language=JavaScript>{alert('发什么神经吗"&fromname&" 到底是替别人清理粪库还是替自己清理阿！');}</script>"
		Response.End
	else
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别 from 用户 where 姓名='"&sjjh_name&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('清理粪库的人不存在，可能有作弊行为，请跟站长反映！');}</script>"	'测4
	Response.End	
end if
rs.close
rs.open "select 粪库,内力,银两,性别 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn
tosex=rs("性别")
Response.Write "<script language=JavaScript>{alert('要记得常清理粪库喔要不然粪库堆到5000点会死喔！');}</script>"
frag="【夜香师】：<font color=red>※"& fromname &"※</font>正在替，<font color=red>※"& sjjh_name &"※</font>清理粪库真是又臭又脏.这年头2万银两真不好赚^_^<font color=red>※"& fromname &"※</font>清理完立刻洗澡去.."
	   	 conn.execute "update 用户 set  银两=银两-20000,粪库=0,内力=内力-1000 where 姓名='"&sjjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000,内力=内力+1000 where 姓名='"&fromname&"'"




set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>【倒夜香消息】</b></font>"&frag	
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

