<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'���Ǵ󷨡�wWw.happyjh.com��
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
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�������Ǵ󷨣�');}</script>"
	Response.End
end if
if sjjh_jhdj<jhdj_xx then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǵ���Ҫ["&jhdj_xx&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>�����Ǵ󷨡�<font color=" & saycolor & ">"+xxdf(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���Ǵ�
function xxdf(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="���ֽ���"  then
Response.Write "<script Language=Javascript>alert('��ʾ��[���ֽ���]������������ֻ�������ͻ��������ϡ����ŵȽ������ɴ�ս�ģ������˵��ڳ����������ɼ�ǿ���������Ǵ󷨵�[���ַ���]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select ����,���� from �û� where ����='" & sjjh_name &"'" ,conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������˲����Բ���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ��ȡ����!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,�ȼ�,����,����,grade from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������˲����Բ���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
if rs("�ȼ�")<jhdj_xscz then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ͵Ϯ����Ա������ѽ������');}</script>"
	Response.End
end if
to1=rs("����")
nlto=rs("����")
rs.close
randomize timer
s=int(rnd*20)+1
nlto=int((nlto/100))*s
r=int(rnd*4)
Select Case r
Case 1
	xxdf="##<bgsound src=wav/xixing.wav loop=1>ʹ��һ�����Ǵ�,��ȡ%%����<font color=red>"& s &"%</font>,����:<font color=red>"& nlto &"</font>��!����С��ȫ����ʧ��!"
Case 2
	xxdf="##<bgsound src=wav/xixing.wav loop=1>����ȡ%%����,û�гɹ����Լ�ʧȥ����<font color=red>"& s &"%</font>,����:<font color=red>"& nlto &"</font>��!%%������Ц,֪���������˰ɣ�"
	conn.execute "update �û� set ����=����-"&nlto&" where ����='" & sjjh_name &"'"
case 3
	rs.open "select ���� from �û� where ����='" & sjjh_name &"'",conn,2,2
	if rs("����")>50000 then
		conn.execute "update �û� set ����=����-50000 where ����='" & sjjh_name &"'"
		xxdf="##<bgsound src=wav/qt.wav loop=1>͵%%������������,ֻ�����´�㻨��<font color=red>5��</font>��,�ӹ��˽�,����ѽ!"
	else
		conn.execute "update �û� set ״̬='��' where ����='" & sjjh_name &"'"
		xxdf="##<bgsound src=wav/oh_no.wav loop=1>�ոհ��ִ���%%�����,�ͱ��ٸ����˷�����,������׽ȥ������"
		call boot(sjjh_name,"��ȡ�����������ߣ�"&to1&",����������η���")
	end if
	rs.close
case else
	xxdf="##<bgsound src=wav/xixing.wav loop=1>ʹ��һ�����Ǵ�,��ȡ%%����<font color=red>"& s &"%</font>,����:"& nlto &"��!����%%��Ѫ��"
	conn.execute "update �û� set ����=����+"&nlto&" where ����='" & sjjh_name &"'"
	conn.execute "update �û� set ����=����-"&nlto&" where ����='" & to1 &"'"
End Select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
