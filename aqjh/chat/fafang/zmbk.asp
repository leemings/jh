<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<>10 then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷����ֻ��վ���ſ��Է���');window.close();}</script>"
	Response.End 
end if
Response.Buffer=true
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
diaox="<bgsound src=wav/zt.mid loop=1><font color=green>��վ�����</font>"& aqjh_name &"<font color=#ff0000>���ǵ����˵��ҿ�,��������ݵ�������Ա����,���ţ�1�ڡ����ϣ�1ǧ�򡢻�����1����������50��������</font>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
for i=0 to x
conn.Execute "update �û� set ����=����+100000000 where ���='����' and ����='" & online(i) & "'"
conn.Execute "update �û� set ����=����+10000000 where ���='����' and ����='" & online(i) & "'"
conn.Execute "update �û� set ����=����+1000000 where ���='����' and ����='" & online(i) & "'"
conn.Execute "update �û� set ����=����+500000 where ���='����' and ����='" & online(i) & "'"
next
conn.close
set conn=nothing
says=diaox

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
Response.Redirect "../../ok.asp?id=705"
%>