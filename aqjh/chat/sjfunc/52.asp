<%@ LANGUAGE=VBScript codepage ="936" %>
<%'清屏
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_grade<=9 then
	Response.Write "<script language=JavaScript>{alert('清屏需要9级才可以操作！');}</script>"
	Response.End
end if
says="<font color=#ff0000>【消息】为了保证大家的聊天速度，站长:"&aqjh_name&"执行了清屏操作……</font>"
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="清屏"
towhoway=0
towho="大家"
saycolor="660099"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
