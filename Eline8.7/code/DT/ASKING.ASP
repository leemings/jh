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
if sjjh_grade>=6  then
ask=trim(Request.form("ask"))
reply=trim(Request.form("reply"))
ask=lcase(server.HTMLEncode(ask))
reply=lcase(server.HTMLEncode(reply))
silver=Trim(Request.Form("silver"))
if not isnumeric(silver) then Response.Redirect "../error.asp?id=475"
if trim(Request.form("ask"))="" then Response.Redirect "../error.asp?id=475"
if trim(Request.form("reply"))="" then Response.Redirect "../error.asp?id=475"
silver=clng(silver)
if silver<1 then
silver=1
elseif silver>99999 then
silver=99999
end if
ask=replace(ask,"'","")
ask=replace(ask,"or","")
ask=replace(ask,"\","")
ask=replace(ask,"www","")
ask=replace(ask,"xajh","")
ask=replace(ask,"net","")
reply=replace(reply,"'","")
reply=replace(reply,"or","")
reply=replace(reply,"\","")
reply=replace(reply,"www","")
reply=replace(reply,"xajh","")
reply=replace(reply,"net","")
Application.Lock
Application("sjjh_ask")=ask
Application("sjjh_reply")=reply
Application("sjjh_askuser")=sjjh_name
Application("sjjh_asksilver")=silver
Application.UnLock
says="<bgsound src=chat/wav/chuti.WAV loop=1><b>【出 题】<font color=balck>" & Application("sjjh_ask") & "？ </font>正确答案是什么？[提问人]:<font color=red>〖" & Application("sjjh_askuser")&"〗</font></b><font color=green><b>奖励："&Application("sjjh_asksilver")&"两!</b></font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=replace(says,chr(13),"")
says=replace(says,chr(10),"")
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
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>
<%
else
response.redirect "../error.asp?id=425"
end if
%>
