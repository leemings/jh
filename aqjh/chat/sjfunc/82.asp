<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'�����뿪����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_jhtw then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������ҲҪ["&jhdj_jhtw&"]�����У�');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
'call dianzan(towho)
act=0
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
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/��������" then
	says="<font color=green>������������</font><font color=" & saycolor & ">"+jiarutw()+"</font>"
else
	says="<font color=green>���뿪������</font><font color=" & saycolor & ">"+likaitw()+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������
function jiarutw()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select grade,����,���� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹ���Ա�����Բ����������㲻������!');}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������Ų������뿪�Լ�������');}</script>"
	Response.End
end if
if rs("����")<1000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������������Ҫ��������100��');}</script>"
	Response.End
end if
conn.execute "update �û� set ����='����',����='ɱ��',grade=1,����=False where ����='" & aqjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("����")
jiarutw="##���ļ��������������µ�һɱ�֣�$$bluebע�����������󲻿��Ա��������Խ��ܱ��˵�ɱ�����󣬳�Ϊɱ��!$$b"
end function

'�뿪����
function likaitw()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")<>"����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������ɱ��,�����Բ���!');}</script>"
	Response.End
end if
conn.execute "update �û� set ����='����',����='����',grade=1,����=True,���=int(���/2),����=int(����/2) where ����='" & aqjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("����")
likaitw="##�Ѿ�����˴��ɱɱ���������Ϊ����Ч����$$red�ó��Լ������һ�������֯���������뿪������!$$"
end function

sub updatemd(jhmp)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select id,����,����ʱ��,��Ա�ȼ�,����,����,����ͷ��,�Ա�,��������,��ż from �û� where ����='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("����")
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=aqjh_name then
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end sub
%>