<%
'添加名单
sub addnpc (zdnpc_name)
    Application.Lock
    js=1
    onlinelist=Application("aqjh_onlinelist"&nowinroom)
    onlineno=ubound(onlinelist)
    yjl=0
    for i=1 to onlineno
    onuser=split(onlinelist(i),"|")
    if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
    else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
    end if
    next 
    if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
    end if
    Application("aqjh_onlinelist"&nowinroom)=newonlinelist
    Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)&" "&zdnpc_name&" "
    Application("aqjh_npc")=Application("aqjh_npc")&";"&zdnpc_name&"|"
    erase  newonlinelist
    erase  onlinelist
    Application.UnLock
    Session("SayCount")=Application("SayCount")
end sub
'删除npc名单
sub delnpc (nl_name)
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),nl_name& "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " &nl_name& " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&nl_name&"|","")
Application.UnLock
end sub
'消息
Sub showchat(mess)
says=mess   '聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","&session("nowinroom")&");parent.f2.document.af.npc.value='"&Application("aqjh_npc")&"';parent.m.location.reload();<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
j="SayStr"&YuShu1(Application("SayCount"))
Application(j)=Str
Application.UnLock()
End Sub
%>