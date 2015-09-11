<%
'网址提交判断
sub ChkPost()
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		Response.Write "<script Language=Javascript>alert('提示：请不要在站外提交数据！');window.close();</script>"
		Response.End
	end if
end sub
'聊天室发话模块
sub lts(actz,towho,saysj,gs,room)
'actz:发言模式，towho:受话对象，saysj：发言内容，gs:公聊还是私聊，room：房间号
	toname=towho
	if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(towho)&" ")=0 then
		toname="大家"
	end if
	Application.Lock
	says=saysj
	says=replace(says,chr(39),"\'")
	says=replace(says,chr(34),"\"&chr(34))
	act=actz
	towhoway=gs
	towho=toname
	addwordcolor="660099"
	saycolor="008888"
	addsays="对"
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& room & ");<"&"/script>"
	addmsg saystr
end sub
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