<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�¶���wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	call mess ("��ʾ����["&chatinfo(0)&"]���䲻�����¶���",1)
end if
if Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<21) and chatinfo(0)="����E��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ƕᱦʱ�䣬����ʹ�öᱦ�书��');}</script>"
	Response.End 
end if
if sjjh_jhdj<=jhdj_duyao then
	call mess ("��ʾ���¶���Ҫ["&jhdj_duyao&"]���ſ��Բ�����",1)
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>���¶���<font color=" & saycolor & ">"+xiadu(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
'�¶�
function xiadu(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	call mess ("��ʾ�����ɣ�������ʲô���뵷���𣿣�",1)
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
	call mess ("��ʾ���������󣬸�ʽ���£�[��Ʒ��&����]",1)
end if
zt=split(fn1,"&")
if not isnumeric(zt(1)) then 
	call mess ("��ʾ����������������ʹ������!",1)
end if
zswupin=zt(0)
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	call mess ("��ʾ����Ʒ����Ӧ����0С��100��",1)
end if
erase zt
'�ж�ɱ����
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="E�߽���"  then
Response.Write "<script Language=Javascript>alert('��ʾ��[E�߽���]������������ֻ�������ͻ��������ϡ����ŵȽ������ɴ�ս�ģ������˵��ڳ����������ɼ�ǿ�����¶���[����E��]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select ����,����,�ȼ�,����,grade,����ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǳ����˲����Բ�����",1)
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and sjjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ��"&to1&"�ձ�����ɱ�����������ɡ���",1)
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���벻Ҫ�Թٸ����й�������",1)
end if
if rs("�ȼ�")<=2 and rs("����")="��" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���벻Ҫ�Գ��뽭�����ֲ�����",1)
end if
if rs("����")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���Է��������������벻Ҫ͵Ϯ!",1)
end if
rs.close
rs.open "select ����,����,ɱ����,grade,w2,����ʱ��,���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǳ����˲����Բ�����",1)
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ����ոձ�����ɱ����������������˵��",1)
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǹٸ������벻Ҫ���й�����",1)
end if
if rs("ɱ����")>=int(Application("sjjh_killman")) and sjjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ��������̫�࣬��������ɱ������"& Application("sjjh_killman") &"�������¶��ˣ�",1)
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����������������벻Ҫ�¶�!",1)
end if
'ɾ���Լ��Ķ�ҩ
duyao=abate(rs("w2"),zswupin,wusl)
conn.execute "update �û� set w2='"&duyao&"' where  ����='"&sjjh_name&"'"
rs.close
'ȡ����ҩɱ����
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�������Ʒ["&zswupin&"]��ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����",1)
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
'�û�
conn.execute "update �û� set ����=����-20,����=����-" & int(tl/4) & " where ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����-" & nl & ",����=����-" & tl & " where ����='" & to1 &"'"
rs.close
over=""
rs.open "select ���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|�¶�:��С���Լ��ж��ˡ�',����ʱ��=now() where ����='" & sjjh_name &"'"
	over="##��С���Լ�Ҳ���˾綾������һ���������ˡ�"
	call boot(sjjh_name,"�¶�����С���Լ��ж��ˣ�")
end if
rs.close
rs.open "select ����,����,���� from �û� where ����='" & to1 &"'",conn,2,2
to1=rs("����")
if rs("����")>-100 then
	xiadu="##<bgsound src=wav/anqi.wav loop=1>��%%��["&zswupin&"]��:"&wusl&"��,ʹ%%������:-<font color=red>" & nl & "</font>��,����:-<font color=red>" & tl & "</font>��!������Ⱦ"& sjjh_name &"����:-<font color=red>"& int((tl)/4 ) &"</font>��!"&over
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
	xiadu="##��%%�ı���:"& Application("sjjh_baowuname") &"���ߣ���õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
	xiadu=sjjh_name & "<bgsound src=wav/daipu.wav loop=1>��" & to1 & "͵͵���˶�ҩ["&zswupin&"]��:"&wusl&"��,ʹ" & to1 & "������:-" & nl & "����:-" & tl & "," & to1 & "����ǰ˵�����������,Ϊ�ұ��𡭡���"&over
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|�¶�:"&zswupin&"',����ʱ��=now() where ����='" & to1 &"'"
'��¼
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"�¶��������ߣ�"&sjjh_name&","&zswupin)
end function
%>