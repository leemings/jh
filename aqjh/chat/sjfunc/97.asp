<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'С��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��С��������<font color=" & sayscolor & ">"+xiaohai(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xiaohai(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select boy from �û� where ����='"&aqjh_name&"'",conn,3,3
if  rs("boy")<>"" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���ƻ�����,ֻ��һ��ѽ��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select �Ա�,��ż from �û� where ����='"&aqjh_name&"'",conn,3,3
if rs("�Ա�")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������е�Ҳ����ѽ��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if isnull(rs("��ż")) or rs("��ż")="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û�У����Լ�һ������ѽ��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select ����,���,����,���� from �û� where ����='"&aqjh_name&"'",conn,3,3
if rs("����")<1000000 and rs("���")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������������100����߽��û��10����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if rs("����")<10000 and rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������������1����ߵ��²���1ǧ��');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
rs.open "select ��������,��ż from �û� where ����='"&aqjh_name&"'",conn,3,3
if DateDiff("d",rs("��������"),date())<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ŀ��ֻ������ȶ�����2���ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zz=rs("��ż")
huaname=trim(fn1)
yin=1000000
jinbi=10
tili=10000
dd=1000
randomize
tt=(int(rnd()*999)+51)
if tt mod 2=0 then
   huasex="��"
   boysex="images/boy.gif"
else
   huasex="Ů"
   boysex="images/girl.gif"
end if
zstemp=huaname&"|"&huasex&"|"&now()&"|1000|1000|1000|0"&"|"&now()
'����|��|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29
conn.execute "update �û� set ����=����-" & yin & ",���=���-"&jinbi&",����=����-"&tili&",����ʱ��=now(),boy='"&zstemp&"',boysex='"&boysex&"' where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����ʱ��=now(),boy='"&zstemp&"',boysex='"&boysex&"' where ����='"&zz&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& zz &"','����','���˸��Ա���')"
xiaohai="##���û���̫��į������һ��˼�붷����##��<font color=red>"&zz&"</font>���װ�������[����]��������������׼��������һ��<font color=red>"&huasex&"Ӥ</font>�����ֽ�<font color=red>"&huaname&"</font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>