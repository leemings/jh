<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�ᱦ�¶���wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<=jhdj_duyao then
	call mess ("��ʾ���¶���Ҫ["&jhdj_duyao&"]���ſ��Բ�����",1)
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
call roompd("�ᱦ�¶�")
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
says="<font color=green>���ᱦ�¶���<font color=" & saycolor & ">"+xiadu(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
'�ᱦ�¶�
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
if wusl=0 or wusl>30 then
	call mess ("��ʾ����Ʒ����Ӧ����0С��30��",1)
end if
erase zt
toname=to1
'�ж�ɱ����
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,����,�ȼ�,grade from �û� where ����='" & to1 &"'",conn,2,2
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ���벻Ҫ�Թٸ����й�����",1)
end if
tonameid=rs("id")
rs.close
rs.open "select id,����,grade,w2 from �û� where ����='" & sjjh_name &"'",conn,2,2
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess ("��ʾ�����ǹٸ������벻Ҫ���й�����",1)
end if
mynameid=rs("id")
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
newover=""
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|�¶�:��С���Լ��ж��ˡ�',����ʱ��=now() where ����='" & sjjh_name &"'"
	over="##��С���Լ�Ҳ���˾綾������һ���������ˡ�"
	newover=sjjh_name&"��С������Ҳ���˾綾������һ���������ˡ�"
	call boot(sjjh_name,"�¶�����С���Լ��ж��ˣ�")
end if
rs.close
rs.open "select ����,���� from �û� where ����='" & to1 &"'",conn,2,2
to1=rs("����")
if rs("����")>-100 then
	xiadu="##<bgsound src=wav/anqi.wav loop=1>��%%��["&zswupin&"]��:"&wusl&"��,ʹ%%������:-<font color=red>" & nl & "</font>��,����:-<font color=red>" & tl & "</font>��!������Ⱦ"& sjjh_name &"����:-<font color=red>"& int((tl)/4 ) &"</font>��!"&over
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	if newover<>"" then
		bbb=zxrsbd(toname,tonameid)
		newxiadu=replace(xiadu,"##",sjjh_name)
		newxiadu=replace(newxiadu,"%%",toname)
		xiadu=xiadu & over & "<br><br>"&bbb
		newover=newxiadu&"<br><br>"&bbb
		call dbxx(newover)
		to1=toname
	end if
	exit function
end if
conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & sjjh_name &"'"
xiadu=sjjh_name & "<bgsound src=wav/daipu.wav loop=1>��" & to1 & "͵͵���˶�ҩ["&zswupin&"]��:"&wusl&"��,ʹ" & to1 & "������:-" & nl & "����:-" & tl & "," & to1 & "����ǰ˵�����������,Ϊ�ұ��𡭡���"&over
conn.execute "update �û� set ״̬='��',�¼�ԭ��='"&sjjh_name&"|�¶�:"&zswupin&"',����ʱ��=now() where ����='" & to1 &"'"
'��¼
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&zswupin&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(to1,"�¶��������ߣ�"&sjjh_name&","&zswupin)
bbb=zxrsbd(sjjh_name,mynameid)
newover=xiadu&"<BR><BR>"&bbb
xiadu=xiadu&"<BR><BR>"&bbb
call dbxx(newover)
to1=toname
end function
%>