<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'����/���ס�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(0)="����E��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫ���һ��׵���ķ���ȥ���ᱦ�����в����Գ��ң�');}</script>"
	Response.End
end if
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="���ֽ���"  then
Response.Write "<script Language=Javascript>alert('��ʾ��[���ֽ���]������������ֻ�������ͻ��������ϣ����ŵȽ������ɴ�ս�������˵��ڳ����������ɼ�ǿ�����ܵ�[����E��]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
erase sjjh_roominfo
erase chatinfo
if sjjh_jhdj<sjjh_jhcz then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ҲҪ["&sjjh_jhcz&"]�����У�');}</script>"
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
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
if mycz="/����" then
	says="<font color=green>�����ƺ쳾��</font><font color=" & saycolor & ">"+chujia()+"</font>"
else
	says="<font color=green>�����ס�</font><font color=" & saycolor & ">"+huanshu()+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function chujia()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")

rs.open "select grade,���,��ż,����,���� from �û� where ����='" & sjjh_name &"'",conn,2,2
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
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������Գ���!');}</script>"
	Response.End
end if
if rs("���")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������Ų������뿪�Լ�������');}</script>"
	Response.End
end if
if rs("��ż")<>"��" or rs("����")<>"��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����쳾����δ��,���ܳ���\n(��������ż������)��');}</script>"
	Response.End
end if
conn.execute "update �û� set ����='����',���='����',grade=1,����=True where ����='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("����")
chujia="##���ƺ쳾,�������κεľ���,���������䷢����,�Ժ󽭺��ϵĴ��ɱɱ,������Թ�����޹�~~"
end function

'����
function huanshu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ���� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("����")<>"����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ǳ�����,�����Բ���!');}</script>"
	Response.End
end if
conn.execute "update �û� set ����='����',���='����',grade=1,����=True,���=int(���/2),����=int(����/2) where ����='" & sjjh_name &"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
updatemd("����")
huanshu="##�쳾����Ϊ��,���첻һ�����,���Ķ���,���ջ��Ǿ����뿪�˵�,��ʱ���Լ�����Ǯ������һ����������ֵ�~~"
end function

sub updatemd(jhmp)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select id,����,����ʱ��,��Ա�ȼ�,���,����,����ͷ��,�Ա�,��������,��ż from �û� where ����='" & sjjh_name &"'",conn,2,2
sjjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("���")
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=sjjh_name then
		sjjh_zm=sjjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&sjjh_jhdj&"|"&sjjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)
		onlinelist(i)=sjjh_zm
		exit for
	end if
next
Application("sjjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end sub
%>