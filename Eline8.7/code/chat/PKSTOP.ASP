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
if sjjh_grade<6  then
	Response.Write "<script Language=Javascript>alert('提示：不是官府中人，不能以操作!');window.close();</script>"
	response.end
end if
f=Minute(time())
if f<30 then
	Response.Write "<script language=JavaScript>{alert('动武时间为每小时的后30分钟！');window.close();}</script>"
	Response.End 
end if
Application.Lock
Application("sjjh_dalie")="PK" 
Application.UnLock
says="<div align='center'><font size=2 color=#FF0000 ><img src='img/bad.gif' border='0'>注意:PK时间结束,停止PK,单挑请在PK时间进行,请大家跟聊管合作,违者抓!<img src='img/bad.gif' border='0'></font></div>"			'聊天数据
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
Response.Write "<script Language=Javascript>alert('提示：执行PK结束操作成功');window.close();</script>"
response.end
%>
