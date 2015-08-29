<%@ LANGUAGE=VBScript codepage ="936" %><%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade>=7 then 
Response.Buffer=true
youname=sjjh_name
useronlinename=Application("sjjh_useronlinename"&nowinroom)

online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
yy="<bgsound src=zt.mid loop=1>"
randomize timer
s=int(rnd*130)
diaox="<font color=#000000>路经此地，特为聊天室里的各大小虾传授"& s &"体力</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
for i=0 to x
conn.Execute "update 用户 set 体力=体力+" & s & " where 姓名='" & online(i) & "'"
next
conn.close
set conn=nothing
says="<font color=#ff0000>【体力】" & youname & ""& diaox &""& yy &"</font>"		'聊天数据

says=replace(says,chr(39),"\'")
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
Response.Write "<script language=javascript>alert('完成！');history.back();</script>"
end if%>
