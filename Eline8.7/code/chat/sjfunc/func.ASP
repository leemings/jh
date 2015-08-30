<%
'踢人模块♀wWw.happyjh.com♀
sub boot(to1,czsj)
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("sjjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),to1 & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
Application("sjjh_useronlinename"&nowinroom)=Replace(Application("sjjh_useronlinename"&nowinroom)," " & to1 & " ","")
Application.UnLock
towhoway=0
saycolor="660099"
addsays="对"
says1="<b>【江湖消息】</b>"&czsj
says1=replace(says1,"'","\'")
says1=replace(says1,chr(34),"\"&chr(34))
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & "踢出" & chr(39) &","& chr(39) & to1 & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & to1 & chr(39) &"," & chr(39) & says1 & chr(39) &",0," & session("nowinroom") & ");</script>"
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