<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�ᱦͶ����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
call roompd("�ᱦͶ��")
if sjjh_jhdj<=jhdj_aq then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ͷ����Ҫ["&jhdj_aq&"]���ſ��Բ�����');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
if towho=sjjh_name or towho="���" or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Զ����ѻ�Ի����ˡ����ʹ�ã�');}</script>"
	Response.End
end if
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
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
says="<font color=green>���ᱦͶ����<font color=" & saycolor & ">"+touzi(mid(says,i+1),towho)+"</font>"
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
if wusl=0 or wusl>30 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��30��');}</script>"
	Response.End 
end if
erase zt
'�ж�ɱ����
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,grade from �û� where ����='" & to1 &"'",conn,2,2
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
tonameid=rs("id")
rs.close
rs.open "select id,����,����,ɱ����,grade,w4,����ʱ�� from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("grade")>=6 and sjjh_name<>application("sjjh_user") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if
mynameid=rs("id")
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
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ʒ["&zswupin&"]��ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա��');}</script>"
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
	newover=sjjh_name&"ѧ�ղ��������������Լ�������һ���������ˡ�"
	call boot(sjjh_name,"��������С�������Լ���")
end if
rs.close
rs.open "select ����,���� from �û� where ����='" & to1 &"'",conn,2,2
to1=rs("����")
if rs("����")>-100 then
	touzi="##<bgsound src=wav/anqi.wav loop=1>��%%��["&zswupin&"]��:"&wusl&"��,ʹ%%��������:<font color=red>-" & nl & "</font>������:<font color=red>-" & tl & "</font>!�Լ�ȴ��������:<font color=red>-"& int(nl/4 ) &"</font>��!"&over
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if newover<>"" then
		bbb=zxrsbd(toname,tonameid)
		newtouzi=replace(touzi,"##",sjjh_name)
		newxiadu=replace(newtouzi,"%%",toname)
		touzi=touzi & over & "<br><br>"&bbb
		newover=newxiadu&"<br><br>"&bbb
		call dbxx(newover)
		to1=toname
	end if
	exit function
end if
conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & sjjh_name &"'"
touzi="##<bgsound src=wav/daipu.wav loop=1>��%%Ͷ����["&zswupin&"]��:"&wusl&"��,ʹ%%������::<font color=red>-" & nl & "</font>����::<font color=red>-" & tl & "</font>," & to1 & "����ǰ˵�����������,Ϊ�ұ��𡭡���"&over
conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|����:"&zswupin&"' where ����='" & to1 &"'"
'��¼
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"�����������ߣ�"&sjjh_name&","&zswupin)
bbb=zxrsbd(sjjh_name,mynameid)
newover=touzi&"<BR><BR>"&bbb
touzi=touzi&"<BR><BR>"&bbb
call dbxx(newover)
to1=toname
end function
%>