<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'ͬ���ھ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ��ͬ���ھ����ܣ�');}</script>"
	Response.End
end if
if aqjh_jhdj<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��ͬ���ھ���Ҫ10�����ϲſ��Բ�����');}</script>"
	Response.End
end if
if aqjh_grade>=6 then
	Response.Write "<script language=JavaScript>{alert('��Ϊ�ٸ���Ա������˴�ܣ���');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if towho="���" or towho=aqjh_name or towho=application("aqjh_automanname") or towho="" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��ͬ���ھ���<font color=" & saycolor & ">"+tgyj(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tgyj(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,����,�书,����,�ȼ�,grade,���� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Է����������������㲻����������ɣ�');}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ǳ����ˣ����ѿ��ƺ쳾������һ�ж�Թ����������޹��ˣ�');}</script>"
	Response.End
end if
if rs("�ȼ�")<30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻�������һ������ͬ���뾡�ɣ�');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻���ԶԹٸ���Աʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
if aqjh_jhdj-rs("�ȼ�")>15 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ս���ȼ��߳��Է�15�������ˣ��㻹�򲻹���������Ҫͬ���ھ�ѽ��');}</script>"
	Response.End
end if
totl=rs("����")
tonl=rs("����")
towg=rs("�书")
tofy=rs("����")
rs.close
rs.open "select ɱ����,����,����,����,����,�书,����,����,����,���,�ȼ�,ͨ�� from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ǳ����ˣ�����һ�ж�Թ����������޹��ˣ�');}</script>"
	Response.End
end if
if rs("ͨ��")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����ͨ������������ʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
if rs("����")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('������������Ҫ��ͱ���ͬ���ھ��������ȳ��أ�');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) and aqjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"����ʹ�ô˹��ܣ�');}</script>"
	response.end
end if
mytl=rs("����")
mynl=rs("����")
mywg=rs("�书")
if mytl<5000 or mynl<3000 or mywg<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫ��ͱ���ͬ���ھ��������������5000����������3000���书����1000��');}</script>"
	response.end
end if
myfy=rs("����")
killer=mytl+mynl+mywg+myfy
yingliang=int(rs("����")/500)

randomize timer
r=int(rnd*1999)+1
killer=killer+yingliang+r*aqjh_jhdj-towg-tofy
if killer<=4000 then
	killer=int(rnd*3999)+1
end if
tlkill=int(killer*2/3)
nlkill=killer-tlkill
shengtl=totl-tlkill
shengnl=tonl-nlkill
shengwg=towg-mywg
if shengwg<0 then shengwg=0
if shengtl<=-100 then
	conn.execute "update �û� set ����=-100,����=-100,�书=0,����=0,����=0,״̬='��',ɱ����=ɱ����+1,����='��',��������=0,�¼�ԭ��='"& aqjh_name & "|��"&to1&"ͬ���뾡' where ����='"& aqjh_name &"'"
	conn.execute "update �û� set ����=" & shengtl & ",����=" & shengnl & ",�书="& shengwg &",״̬='��',����='��',��������=0,�¼�ԭ��='"& aqjh_name &"|����ͬ���뾡' where ����='" & to1 & "'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'" & aqjh_name & "','ͬ���뾡','����')"
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','ͬ���뾡','����')"
	e="<br>%%Ҳ������##ȥ�������ȥ�ˡ�"
	call boot(to1,aqjh_name&"����ͬ���ھ�")
else
	conn.execute "update �û� set ����=-100,����=-100,�书=0,����=0,����=0,״̬='��',����='��',��������=0,�¼�ԭ��='"&aqjh_name&"|��"&to1&"ͬ���뾡' where ����='"& aqjh_name &"'"
	conn.execute "update �û� set ����=" & shengtl & ",����=" & shengnl & ",�书="& shengwg &" where ����='" & to1 & "'"
	e=""
end if
tgyj="##<bgsound src=wav/tgyj.wav loop=1>����һ����Х��ȫ��������ɫ��������Χ��ͻȻ�����ź�������һ��գ�۾ͷ���%%����������ʱ����һ�������ۺ�⣬##ȫ���Ѫ�ⶼ����һ����������%%��ȥ��%%㲲���������ɱ��������" & tlkill & "��������"& nlkill &"���书�����½�" & mywg & "�㡣##����ǰ������ź��þûص������������" & e
call boot(aqjh_name,"��"& to1 &"ͬ���ھ�")
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>