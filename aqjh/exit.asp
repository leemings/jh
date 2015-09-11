<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name<>"" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
aqjh_name=aqjh_name
conn.Execute "update 用户 set 宝物='无' where  姓名='"&aqjh_name&"'"
conn.close
set conn=nothing
end if
if Session("aqjh_inthechat")="1" and aqjh_name<>"" then
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),aqjh_name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & aqjh_name & " ","")
Application.UnLock
'****************
says="<font color=black>【公告】</font><font color=F08000>" & Replace(Application("aqjh_userout"),"%%","<font color=black>" & aqjh_name & "</font>") & "</font><font class=t>(" & t & ")</font>"			'聊天数据
says=replace(says,"'","\'")
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
end if
Session("aqjh_inthechat")="0"
Session.Abandon
Response.Redirect "index.asp"
response.end
%>
