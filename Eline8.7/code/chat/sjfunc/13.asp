<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'Ͷ����wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����Ͷ����');}</script>"
	Response.End
end if
if Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<21) and chatinfo(0)="����E��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ƕᱦʱ�䣬����ʹ�öᱦ�书��');}</script>"
	Response.End 
end if
if sjjh_jhdj<=jhdj_aq then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ͷ����Ҫ["&jhdj_aq&"]���ſ��Բ�����');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��Ͷ����<font color=" & saycolor & ">"+touzi(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)

'Ͷ��
function touzi(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('��ʾ���������󣬸�ʽ���£�[��Ʒ��&����]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ʹ�����֣�');}</script>"
	Response.End 
end if
zswupin=zt(0)
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��100��');}</script>"
	Response.End 
end if
erase zt
'�ж�ɱ����
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="E�߽���"  then
Response.Write "<script Language=Javascript>alert('��ʾ��[E�߽���]������������ֻ�������ͻ��������ϡ����ŵȽ������ɴ�ս�ģ������˵��ڳ����������ɼ�ǿ�����ð�����[����E��]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select ����,����,�ȼ�,����,grade,����ʱ��,���� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ����������������˵��');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
if rs("�ȼ�")<=15 and rs("����")="��"  and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
if rs("����")=True   and sjjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,ɱ����,grade,w4,����ʱ�� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("grade")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոձ�����ɱ����������������˵��');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("sjjh_killman"))   and sjjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & sjjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('������̫�࣬��������ɱ������"& Application("sjjh_killman") &"������Ͷ���ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���������������벻ҪͶ��!');}</script>"
	Response.End
end if
'ɾ���Լ��İ���
duyao=abate(rs("w4"),zswupin,wusl)
conn.execute "update �û� set w4='"&duyao&"' where  ����='"&sjjh_name&"'"
rs.close
'ȡ������ɱ����
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ʒ["&zswupin&"]��ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'�û�
conn.execute "update �û� set ����=����-20,����=����-" & int(nl/4) & " where ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����-" & nl & ",����=����-" & tl & " where ����='" & to1 &"'"
rs.close
over=""
rs.open "select ���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|����:��С�İ��������Լ���',����ʱ��=now() where ����='" & sjjh_name &"'"
	over="##ѧ�ղ��������������Լ�������һ���������ˡ�"
	call boot(sjjh_name,"��������С�������Լ���")
end if
rs.close
rs.open "select ����,����,���� from �û� where ����='" & to1 &"'",conn,2,2
to1=rs("����")
if rs("����")>-100 then
	touzi="##<bgsound src=wav/anqi.wav loop=1>��%%��["&zswupin&"]��:"&wusl&"��,ʹ%%��������:<font color=red>-" & nl & "</font>������:<font color=red>-" & tl & "</font>!�Լ�ȴ��������:<font color=red>-"& int(nl/4 ) &"</font>��!"&over
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & sjjh_name &"'"
if rs("����")=Application("sjjh_baowuname") then
	conn.execute "update �û� set ��������=0,����='"& Application("sjjh_baowuname") &"' where ����='" & sjjh_name &"'"
	conn.execute "update �û� set ��������=0,����='��',����=true,����=100,����=2000 where ����='" & to1 &"'"
	touzi="##��%%�ı���:"& Application("sjjh_baowuname") &"���ߣ���õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
touzi="##<bgsound src=wav/daipu.wav loop=1>��%%Ͷ����["&zswupin&"]��:"&wusl&"��,ʹ%%������::<font color=red>-" & nl & "</font>����::<font color=red>-" & tl & "</font>," & to1 & "����ǰ˵�����������,Ϊ�ұ��𡭡���"&over
conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|����:"&zswupin&"' where ����='" & to1 &"'"
'��¼
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"�����������ߣ�"&sjjh_name&","&zswupin)
end function
%>