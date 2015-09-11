<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade=10 then 
Response.Buffer=true
youname=aqjh_name
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
yy="<bgsound src=wav/lw.mid loop=1>"
diaox="<font color=#ff00ff><img src='img/1-2.gif' border='0'>站长发级了，每位在线的玩家全部升一级,不是骗你的，等会还会发，云“钟声是我的问候，歌声是我的祝福，雪花是我的贺卡，美酒是我的飞吻，清风是我的拥抱，快乐是我的礼物！统统都送给你，my love!!”</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")

aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for j=0 to chatroomnum	
	useronlinename=Application("aqjh_useronlinename"&j)
	online=Split(Trim(useronlinename)," ",-1)
	x=UBound(online)
	for i=0 to x
		rs.open "select allvalue from 用户 where 姓名='" & online(i) & "'",conn
		if not (rs.EOF and rs.BOF) then
		jhjy=rs("allvalue")
		jhdj=int(sqr(jhjy/50))
		jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
		jhdj=jhdj+1
		jhjy=jhjy+jhadd
		conn.Execute ("update 用户 set allvalue="&jhjy&",等级="&jhdj&" where 姓名='" & online(i) & "'")
		end if
		rs.close
	next
next
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【站长发级】" & youname & ""& diaox &""& yy &"</font>"		'聊天数据
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
end if%>