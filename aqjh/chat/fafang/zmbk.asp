<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<>10 then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？只有站长才可以发！');window.close();}</script>"
	Response.End 
end if
Response.Buffer=true
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
diaox="<bgsound src=wav/zt.mid loop=1><font color=green>【站长拨款】</font>"& aqjh_name &"<font color=#ff0000>考虑到拉人的幸苦,特向有身份的在线人员拨款,掌门：1亿、长老：1千万、护法：1百万、堂主：50万银两！</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
for i=0 to x
conn.Execute "update 用户 set 银两=银两+100000000 where 身份='掌门' and 姓名='" & online(i) & "'"
conn.Execute "update 用户 set 银两=银两+10000000 where 身份='长老' and 姓名='" & online(i) & "'"
conn.Execute "update 用户 set 银两=银两+1000000 where 身份='护法' and 姓名='" & online(i) & "'"
conn.Execute "update 用户 set 银两=银两+500000 where 身份='堂主' and 姓名='" & online(i) & "'"
next
conn.close
set conn=nothing
says=diaox

says=replace(says,chr(39),"\'")
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
Response.Redirect "../../ok.asp?id=705"
%>