<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("advtime")<>"" then
if session("advtime")>now()-0.0006 then  Response.Redirect "../error.asp?id=490"
end if
session("advtime")=now()
if sjjh_grade<6  then
	Response.Write "<script Language=Javascript>alert('提示：不是官府中人，不能以操作!');window.close();</script>"
	response.end
end if
if Application("sjjh_dalie")="土匪" then
	Response.Write "<script Language=Javascript>alert('提示：土匪还在呢！');window.close();</script>"
	response.end
end if
Application.Lock
Application("sjjh_dalie")="土匪" 
Application.UnLock
says="<bgsound src=wav/fei.wav loop=3><font color=red>【消息】</font><b><font color=red>突然一群土匪<img src=img/jiao.gif>闯入江湖抢劫，请高手们快去剿匪啊。</font></b>"			'聊天数据
says=replace(says,"'","\'")
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
Response.Write "<script Language=Javascript>alert('提示：官府放猎物成功！');window.close();</script>"
response.end
%>

