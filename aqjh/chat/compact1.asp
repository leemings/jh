<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<%Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
f=Minute(time()) 
useronlinename=Application("aqjh_useronlinename"&nowinroom)
at=request.form("at")
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
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в�������ѡ�񹥻�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܶ�NPCʹ�ô˲�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
from1=aqjh_name
compact=""
compact1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
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
if rs("�ȼ�")<=50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���޺��弼��Ҫս���ȼ�50�����ϲſ��Բ�����');location.href = 'javascript:history.go(-1)';</script>"
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
peiou=rs("��ż")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(PEIOU)&" ")=0 Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����ż["&peiou&"]û�����������в��ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn
zstt=rs("ת��")
if zstt>=5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է���ת��5���ˣ������˲��˶Է���');}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
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
jhhy=rs("��Ա�ȼ�")
ntnt=rs("�ȼ�")
if rs("�ȼ�")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('[to1]Ϊ���ٽ������֣�����һ��ô�ص���ʽ�ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "SELECT * FROM t WHERE b='" & aqjh_name & "' or c='" & aqjh_name & "' and a='" & at & " '",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�㲢û��["&at&"]�����ķ��޺��弼ѽ��');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
nei=abs(rs("d"))
rs.close
rs.open "select * from �û� where ����='" & peiou & "'",conn
	htwg1=rs("�书")
	htgj1=rs("����")
rs.close	
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
if rs("����")<nei then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�������������������кϺμ���');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
htwg2=rs("�书")
htgj2=rs("����")
rs.close
rs.open "select * from �û� where ����='" & to1 & "'",conn
towg=rs("�书")
tofy=rs("����")
killer=(((htgj1+htgj2)-tofy)+nei)/100
'���ɱ����С��100�����
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
conn.execute "update �û� set ����=����-" & nei & ",����ʱ��=now() where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 & "'"
e=""
if rs("����")<-100 then
	conn.execute "update �û� set ɱ����=ɱ����+1 where ����='" & aqjh_name &"'"
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&aqjh_name&"|���弼:"&at&"' where ����='" & to1 & "'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & from1 & "','���弼','����')"
	e="�㣬" & to1 & "�����ĵ�����ȥ�����Ӵ˽�����������һֻ��Ϻ"
	compact="["&from1 & "]����" & nei & "��������{" & peiou & "}һ���(" & to1 & ")ʹ������Ϊ" & at & "�ķ��޺��弼��ɱ��" & killer & e
	call boot(to1,"���弼�������ߣ�"&aqjh_name&","&at)	
else
	compact="["&from1 & "]����" & nei & "��������{" & peiou & "}һ���(" & to1 & ")ʹ������Ϊ" & at & "�ķ��޺��弼��ɱ��" & killer & e
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>�����弼��"& compact &"</font>"& t			'��������
says=replace(says,"'","\'")
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
Response.Write "<script Language=Javascript>alert('��ϲ�����ķ��޺��弼�Ѿ���ɣ�');location.href = 'javascript:history.go(-1)';</script>"
%>