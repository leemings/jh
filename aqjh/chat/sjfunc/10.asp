<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'�¶�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<=jhdj_duyao then
	call mess ("��ʾ���¶���Ҫ["&jhdj_duyao&"]���ſ��Բ�����",1)
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
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
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
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,�ȼ�,����,grade,����ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǳ����˲����Բ�����",1)
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and aqjh_grade<10  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ��"&to1&"�ձ�����ɱ�����������ɡ���",1)
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���벻Ҫ�Թٸ����й�������",1)
end if
if rs("�ȼ�")<=30 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���벻Ҫ�Գ��뽭�����ֲ�����",1)
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���Է��������������벻Ҫ͵Ϯ!",1)
end if
rs.close
rs.open "select ����,����,ɱ����,grade,w2,����ʱ��,����,ɱ��ʱ�� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǳ����˲����Բ�����",1)
end if
sj=DateDiff("n",rs("ɱ��ʱ��"),now())
if sj<30 then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        call mess ("��ʾ����ո�ɱ��һ���ˣ���30���Ӻ��ٿ���ɱ���ˣ�",1)
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 then
		rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ����ոձ�����ɱ����������������˵��",1)
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǹٸ������벻Ҫ���й�����",1)
end if
if rs("ɱ����")>=int(Application("aqjh_killman")) and aqjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�������¶��ˣ�",1)
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
conn.execute "update �û� set w2='"&duyao&"' where  ����='"&aqjh_name&"'"
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
conn.execute "update �û� set ����=����-20,����=����-" & int(tl/4) & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-" & nl & ",����=����-" & tl & " where ����='" & to1 &"'"
rs.close
over=""
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&aqjh_name&"|�¶�:��С���Լ��ж��ˡ�',����ʱ��=now() where ����='" & aqjh_name &"'"
	over="##��С���Լ�Ҳ���˾綾������һ���������ˡ�"
	call boot(aqjh_name,"�¶�����С���Լ��ж��ˣ�")
end if
rs.close
rs.open "select ����,����,����,ɱ��ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
to1=rs("����")
if rs("����")>-100 then
	xiadu="##<bgsound src=wav/anqi.wav loop=1>��%%��["&zswupin&"]��:"&wusl&"��,ʹ%%������:-<font color=red>" & nl & "</font>��,����:-<font color=red>" & tl & "</font>��!������Ⱦ"& aqjh_name &"����:-<font color=red>"& int((tl)/4 ) &"</font>��!"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function	
end if
conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name &"'"
if rs("����")=Application("aqjh_baowuname") then
	conn.execute "update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where ����='" & aqjh_name &"'"
	conn.execute "update �û� set ��������=0,����='��',����=true,����=100,����=2000 where ����='" & to1 &"'"
	xiadu="##��%%�ı���:"& Application("aqjh_baowuname") &"���ߣ���õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
	xiadu=aqjh_name & "<bgsound src=wav/daipu.wav loop=1>��" & to1 & "͵͵���˶�ҩ["&zswupin&"]��:"&wusl&"��,ʹ" & to1 & "������:-" & nl & "����:-" & tl & "," & to1 & "����ǰ˵�����������,Ϊ�ұ��𡭡���"&over
	conn.execute "update �û� set ɱ��ʱ��=now(),״̬='��',�¼�ԭ��='"&aqjh_name&"|�¶�:"&zswupin&"',����ʱ��=now() where ����='" & to1 &"'"
'��¼
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','"&zswupin&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"�¶��������ߣ�"&aqjh_name&","&zswupin)
end function
%>