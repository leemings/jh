<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
chatroomsn=session("nowinroom")	'��ǰ���ڷ������
chatroomname=Application("aqjh_chatroomname"&mychatroomsn)
mm=Application("aqjh_chatroomname"&chatroomsn)
chatroominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(chatroominfo)-1



aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"

nowinroom=session("nowinroom")
online=split((Application("aqjh_useronlinename"&nowinroom))," ")
aqjh_roominfo=split(Application("aqjh_room"),";")
onlinenum=ubound(online)+1

chatroomnum=ubound(aqjh_roominfo)-1

id=trim(Request.QueryString ("id"))
mode=trim(Request.QueryString ("mode"))
for i=0 to chatroomnum	
	ydl=1
	useronlinename=Application("aqjh_useronlinename"&nowinroom)
if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(id)&" ")<>0 then 
Response.Write "<script language=javascript>alert('�����������޼��������ˣ�');</script>"
Response.End 
end if
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(id)&" ")=0 then ydl=0

next


if aqjh_name=""  then
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û���ٻ�����Ȩ��������\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select * from �û� where ����='"&aqjh_name&"'"),conn
myzs=rs("ת��")
zaohuan1=rs("�ٻ���1")
if zaohuan1=""   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û������,��Ҫ���ң�����\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
if zaohuan1="��"   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û������,��Ҫ���ң�����\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
if rs("����")<10000000   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û��1000���ֽ𣡣���\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
if rs("���")<10   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û��10�����,��Ҫ���ң�����\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
if rs("����")<10000   then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û��10000����,��Ҫ���ң�����\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
if aqjh_name=""  or  myzs<2 then
rs.Close
set rs=nothing
conn.Close
set conn=nothing

Response.Write "<script language=JavaScript>{alert('  �Բ���\n  ��û������Ȩ��������\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End
end if
rs.Close
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
rs.Open ("select * from �û� where ����='"&id&"' "),conn
if rs.BOF or rs.EOF then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  ������ޱ��Ƿ������ˣ��������ݳ���~ ��');}</script>"
Response.End 
end if
if rs("״̬")<>"����" then 
xinxi="�������ޱ������ݺ�,����"&aqjh_name&"������300������5����Ҹ���������"
conn.Execute ("update �û� set ����=����-300,���=���-5 where ����='"&aqjh_name&"'")
else
xinxi="������˯���б�����!~������������Ҫ���ͷ���"
end if
if rs.BOF or rs.EOF then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('  �Բ���\n  Ŀǰ�����޳��֣�����\n  �밴 [ȷ��] �� �أ�');}</script>"
Response.End 
end if




randomize timer
qq=int(rnd()*65*myzs)+1
nn=int(rnd()*5500000*myzs)+1

	ii=int(rnd()*1000000*myzs)+1
	iii=int(rnd()*200000*myzs)+1000
	aa=int(rnd()*1500000*myzs)+1
conn.Execute ("update �û� set ״̬='����',����=false,�书="&ii&"+146534,����="&iii&"+1000,����="&aa&"+1000,�ȼ�="&qq&"*"&myzs&",����="&nn&"+1000,����ʱ��=now(),����ʱ��='2015-3-15 20:24:45' where ����='"&id&"'")
conn.Execute ("update �û� set ����=����-10000000,����=����-1000,���=���-5 where ����='"&aqjh_name&"'")

Session("aqjh_inthechat")="1"
Session("aqjh_savetime")=now()
Session("aqjh_lasttime")=sj
myzanli=0
tjrf=rs("ͨ��")
jhmp=rs("����")
dj=rs("�ȼ�")
newuser=rs("times")
aqjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("���")
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
hymd=rs("��������")
mywife=rs("��ż")
tili=rs("����")
nl=rs("����")
wg=rs("�书")

'myonline = id & "|" & sex & "|" & jhmp & "| npc|" &jhtx& "|" & aqjh_jhdj& "| 0 |"&myzanli&"|"&"����"
myonline = id &"|" & sex & "|"&aqjh_name&"|���ٻ���|" &jhtx& "|"& dj &" |0|9|"&myzanli&"|"&"����"
'myonline = id & "|" & sex & "|"&aqjh_name&"| ����|" &jhtx& "| 0 |"&myzanli&"|"&"����"
Application.Lock
'-----------------------------------------------------------------------------------


'����������������������f3��ʾ�������㷨.

dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'���������ֺ���ƴ������
'if yjl=0 and len(onuser(2))<len(jhmp) then			'���������ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(2))>len(jhmp) then			'���������ֳ���,�̵���ǰ
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'�����ֺ���ƴ������
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'�����ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'�����ֳ���,�̵���ǰ
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if

Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)& " " & id  & " "
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
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
says=replace(says,chr(39),"\'") 
says=replace(says,chr(34),"\"&chr(34)) 

into="���ֽ�������<font color=red><b>id:"& aqjh_name &"</b></font>������<font color=blue><b>"& id &"</b></font>���ٻ�����{"&xinxi&"}��������"&tili&"�㣬����"&nl&"�㣬�书"&wg&"�ȼ�"&dj&"��....."

'-------------------------------------------------------------------------------------------
'�����������������ʾ��Ϣ�Ĵ���
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
'says="<font color=red>��npc���롽</font><font color=green>���ֽ�����վ</font><font color=#0000ff>npc</font><img src="&jhtx&" width=60 height=60><s><font color=#0000ff>"& id &"</font></s>��������,��ҸϿ�ɱ��!~�о����Ŷ��"   '��������
says="<font color=black>���ٻ����ޡ�</font><font color=008800>" & Replace(into,"##","<img src="&jhtx&"><a href=javascript:parent.sw(\'[" & id & "]\'); target=f2>" & id &"</a><font color=red><b>id:"& aqjh_id &"</b></font>") & "</font><bgsound src=readonly/cd.mid loop=1>"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & id & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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


'-------------------------------------------------------





%>

