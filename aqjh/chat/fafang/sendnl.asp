<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade>=7 then
Response.Buffer=true
youname=aqjh_name
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
yy="<bgsound src=wav/zt.mid loop=1>"
randomize timer
s=int(rnd*130)
diaox="<font color=#000000>·���˵أ���Ϊ��������ĸ���СϺ����"& s &"����</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
for i=0 to x
	conn.Execute"update �û� set ����=����+" & s & " where ����='" & online(i) & "'"
next
conn.close
set conn=nothing
says="<font color=#ff0000>��������" & youname & ""& diaox &""& yy &"</font>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
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
Response.Write "<script language=javascript>alert('��ɣ�');window.close();</script>"
end if%>
