<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc/func.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
f=Minute(time()) 
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
at1=request.form("at1")
at2=request.form("at2")
at3=request.form("at3")
to1=request.form("to1")
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('�����ڡ����롱״̬�У���ʹ�á��һ����������ܽ����');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"���ڡ����롱״̬�У��벻Ҫ������');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���������˲��������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܶ�NPCʹ�ô˲�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=aqjh_name
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")="����ѵ��Ӫ" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ѵ��Ӫ���˲����Բ�����');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�ٸ�������ʹ��������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���棺���"& s &"���ٷ���,�ɱ����ţ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('����������ɱ�����������ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("�ȼ�")<=21 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��������Ҫս���ȼ�21�����ϲſ��Բ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ȡ�����������ٲ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn,2,2
zstt=rs("ת��")
if zstt>=5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���ת��5���ˣ������˲��˶Է���');}</script>"
	Response.End
end if
if rs("����")="����ѵ��Ӫ" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǳ���ѵ��Ӫ���˲����Բ�����');}</script>"
	Response.End
end if
jhhy=rs("��Ա�ȼ�")
if rs("����")="�ٸ�" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ԶԹٸ��˲�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("����")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[to1]�������������У����ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("�ȼ�")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[to1]Ϊ���ٽ������֣�����һ��ô�ص���ʽ�ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn
mp=rs("����")
rs.close
rs.open "SELECT * FROM y WHERE a='" & at1 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲢û��["&at1&"]�������书ѽ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei1=abs(rs("d"))
wg1=abs(rs("c"))
rs.close
rs.open "SELECT * FROM y WHERE a='" & at2 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲢û��["&at2&"]�������书ѽ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei2=abs(rs("d"))
wg2=abs(rs("c"))
rs.close
rs.open "SELECT * FROM y WHERE a='" & at3 & "' and b='" & mp & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲢û��["&at3&"]�������书ѽ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei3=abs(rs("d"))
wg3=abs(rs("c"))
nei=nei1+nei2+nei3
wg=wg1+wg2+wg3
rs.close
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
mygj=rs("����")
if rs("����")<nei or rs("�书")<wg then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('������������������������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn
tofy=rs("����")
killer=(mygj-tofy)+nei+wg
'killer=int(((lxjwg1+lxjgj1)-(lxjwg2+lxjgj2)+nei/10)/5)
'ɱ��С��100
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update �û� set ����=����-" & nei & ",�书=�书-"&wg&",����ʱ��=now() where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 & "'"
e=""
if rs("����")<-100 then
	conn.execute "update �û� set ɱ����=ɱ����+1 where ����='" & aqjh_name &"'"
	conn.execute "update �û� set ״̬='��,�¼�ԭ��='"&aqjh_name&"|������:"&at1&at2&at3&"'' where ����='" & to1 & "'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','������','����')"
	e="�㣬" & to1 & "�����ĵ�����ȥ�����Ӵ˽�����������һֻ��Ϻ"
	stunt=aqjh_name & "����������" & to1 & "ʹ���ˡ�" & at1 & "+" & at2 & "+" & at3 & "����һ����������������ɱ��" & killer & e
	call boot(to1,"�������������ߣ�"&aqjh_name&","&at1&at2&at3)
else
	stunt=aqjh_name & "����������" & to1 & "ʹ���ˡ�" & at1 & "+" & at2 & "+" & at3 & "����һ����������������ɱ��" & killer & e
end if
says="<font color=red>����������"& stunt &"</font>"& t			'��������

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
Response.Write "<script Language=Javascript>alert('��ϲ�������������Ѿ�������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
%>
