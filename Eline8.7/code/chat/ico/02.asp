<%@ LANGUAGE=VBScript codepage ="936" %>
<%'˵����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
says="<font color=#000000>��˵���ſ���"&sjjh_name&"Ŀǰִ�в鿴�¹���˵��״��</font>"
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="˵��"
towhoway=0
towho="���"
saycolor="660099"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>parent.optfrm.location.href='../ico/03.asp';</script>"
End Sub
%>
