<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'С������
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=Application("aqjh_onlinekill")
if onlinenow<onlinekill and chatinfo(0)<>"��Թ���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ:�������ߵ���"&onlinekill&"�˲��ö��䣡');}</script>"
	Response.End
end if
next
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
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
f=Minute(time()) 
if f>Application("aqjh_pktime") then
        Response.Write "<script language=JavaScript>{alert('��ʾ������PK���ʱ��ΪÿСʱǰ["&Application("aqjh_pktime")&"]�֣������ȥ��Թ��');}</script>"
	response.end
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
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=cwgj(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'С������
function cwgj(fn1,to1)
fn1=trim(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="ͨ�Խ���"  then
Response.Write "<script Language=Javascript>alert('��ʾ��[ͨ�Խ���]������������ֻ�������ͻ��������ϡ����ŵȽ������ɴ�ս�ģ������˵��ڳ����������ɼ�ǿ�����ܵ�[���־�ս]����ȥ�ɣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
fn1=trim(fn1)
rs.open "select boysex,boy,��ż FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
mypic="<img src=/chat/boy/"&rs("boysex")&" width=30 height=30>"
cwnl="�书"
cwgj="����"
cwfy="����"
zz=rs("��ż")
rs.close
rs.open "select ����,�ȼ�,����,grade,����,����ʱ�� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����ͽ" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǽ̻����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<5 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ձ�����ɱ�����ȷŹ����ɣ�');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
if rs("�ȼ�")<=18 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,grade,����,ɱ����,boy from �û� where  ����='"&aqjh_name&"'",conn,2,2
if rs("����")="����ͽ" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǽ̻����˲����Բ�����');}</script>"
	Response.End
end if
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ����!');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�����ٷ����ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("boy")="" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û��С�����Ͽ���һ����');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ���');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if fn1<>zt(0) then
	Response.Write "<script language=JavaScript>{alert('��ʾ������ô���Լ��ĺ��Ӷ��д�');}</script>"
	Response.End
end if
cssj=clng(DateDiff("d",zt(6),now()))
gs=clng(zt(6))
if gs<(80+cssj*5) then
	czzt="����"
elseif gs>((100+cssj*5+50)) then
	czzt="˳��"
else
	czzt="��ͨ"
end if
if czzt="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����鲻�ã���������һ���ܣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if

if clng(zt(3))<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С���������������ܴ���ˣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
if clng(zt(4))<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С���������������ܴ���ˣ�');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	response.end
end if
rs.close
cwsr=clng(DateDiff("d",zt(2),now()))
if cwsr<10 then
	cwgj="<font color=green>��С��������<font color=" & saycolor & ">##��ʹ��С����������%%����ı������ڲ����ڣ���������(˵����С��������������,С��������������5ȡ��������ȥ�����������������Է������������书ֵ��������ƽ������!)</font>"
	exit function
end if
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"-"&clng(zt(3))*2&","&cwgj&"="&cwgj&"-"&clng(zt(4))*2&","&cwfy&"="&cwfy&"-"&clng(zt(5))*2&" where ����='" & to1 &"'")
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"+"&clng(zt(3))&","&cwgj&"="&cwgj&"+"&clng(zt(4))&","&cwfy&"="&cwfy&"+"&clng(zt(5))&",boy='' where ����='" & aqjh_name &"'")
conn.Execute ("update �û� set "&cwnl&"="&cwnl&"+"&clng(zt(3))&","&cwgj&"="&cwgj&"+"&clng(zt(4))&","&cwfy&"="&cwfy&"+"&clng(zt(5))&",boy='' where ����='" & zz &"'")
conn.Execute ("update �û� set ����=����-"&clng(zt(3))*5&" where ����='" & to1 &"'")
rs.open "select ���� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")<-100 then
	call boot(to1,"���У������ߣ�"&aqjh_name&",["&zt(0)&"]����"&fn1)
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','ʹ��"&fn1&"ͬ���ھ�','����')"
	conn.execute "update �û� set ״̬='��',��¼=now(),�¼�ԭ��='" & aqjh_name & "|ʹ��"&fn1&"ͬ���ھ�' where ����='"&to1&"'"
	cwgj="<font color=green>��С��ͬ���ھ���<font color=" & saycolor & ">##��С��<b>"&zt(0)&"</b>]"&mypic&"ʵ�ڿ�����%%�����𹥻�,ʹ%%�����½�:<font color=red>"&clng(zt(3))*5&"</font>��,�����½�:<font color=red>"&clng(zt(4))*2&"</font>��,�����½�:<font color=red>"&clng(zt(5))*2&"</font>��,С���ؼ�����,%%������[<b>"&cwnl&"</b>]:-"&int(clng(zt(3))*2)&"��,%%����������֧���أ���["&zt(0)&"]���������䡭����С��Ҳ�������ȡ�塭��</font>"
else
	cwgj="<font color=green>��С��ͬ���ھ���<font color=" & saycolor & ">##��С��<b>"&zt(0)&"</b>]"&mypic&"ʵ�ڿ�����%%�����𹥻�,ʹ%%�����½�:<font color=red>"&clng(zt(3))*5&"</font>��,�����½�:<font color=red>"&clng(zt(4))*2&"</font>��,�����½�:<font color=red>"&clng(zt(5))*2&"</font>��,С���ؼ�����,%%������[<b>"&cwnl&"</b>]:-"&int(clng(zt(3))*2)&"��,С��Ҳ�������ȡ�塭��</font>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>