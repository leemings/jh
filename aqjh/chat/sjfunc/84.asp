<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��ƭ����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<50 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����з��ĵȼ����ٵ�50�����У�');}</script>"
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
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>����ƭ���С�</font><font color=" & saycolor & ">"+fmrk(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��ƭ����
function fmrk(to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select grade,����,ְҵ,���,���� from �û� where ����='" & aqjh_name &"'",conn
if rs("ְҵ")<>"�з�" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ְҵ�����з������ܲ�����');}</script>"
	Response.End
end if
if rs("����")<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û�д���5ǧ�������ӣ�С�ı�ץס��һ���ӵ�ţ���������壡');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ�ٸ����뷷���˿ڣ��㲻�����ˣ�');}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ϊ�����˻��������ֵ��°ܻ������飬̫�������ˣ�');}</script>"
	Response.End
end if
rs.close
rs.open "select ID,grade,����,�Ա�,��ż,����,���� from �û� where ����='"& to1 &"'",conn
if rs("�Ա�")<>"��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻�ǿ����˰ɣ�Ů����Ҳ������԰��ƭ��');}</script>"
	Response.End
end if
if rs("����")="����" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����~~~~~����������Ҳ��ƭ��');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻����ˣ��ٸ�����Ҳ��ƭ��');}</script>"
	Response.End
end if
toid=rs("id")
peiou=rs("��ż")
if peiou=aqjh_name then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���������ѵ��Ϲ�Ҳ�����ɣ�');}</script>"
	Response.End
end if
%><!--#include file="data1.asp"--><%
sql="select * from ���� where ����ID="&toid
set rs1=connt.execute(sql)
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	connt.close
	set connt=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ۻ��˰ɣ�������ǹ���������ˣ�');}</script>"
	Response.End
end if
rs1.close
randomize timer
r=int(rnd*10)+1
fn1="��ƭ����"
meimao=rs("����")
yinliang=int(rs("����")/2)
select case r
	case 1
		conn.execute "update �û� set ����=����-50000000,����=����-500,����=����-500 where ����='"&aqjh_name&"'"
		conn.execute "update �û� set ����=����+50000000 where ����='" & to1 & "'"
		fmrk=aqjh_name & "��ƭ��û�ɹ�������"& to1 &"����һͨ����æ����100�������������Ǹ������������ˣ������½�500�������½�500��"
	case 2
		conn.execute "update �û� set ����=����+500000,����=����-1000 where ����='" & aqjh_name & "'" 
		sql="insert into ����(����ID,����,��ò��) values ("& toid &",'" & to1 & "'," & meimao & ")"
set rs1=connt.execute(sql)
		fmrk=aqjh_name & "����Ƥ�����Ǻ��ã����ɴ�ƭ�ľͰ�"& to1 &"�����������ˣ��õ��ô���50�򣬵����½�1000��" 
	case 3
		conn.execute "update �û� set ״̬='��',��¼=now()+1/144,�¼�ԭ��='��ƭ�˿ڲ��ɹ���ץ' where ����='" & aqjh_name & "'"
		conn.execute "insert into l(b,c,d,a,e) values ('ϵͳ','" & to1 & "','����',now(),'" & fn1 & "')"
		fmrk=aqjh_name & "������ƭ"& to1 &"��ȴ���ٸ���Ա���֣�ץȥ�����ˡ�10���Ӻ�ſ��Ե�¼��"
		call boot(aqjh_name,"���Σ������ߣ�"&aqjh_name&","&fn1)
	case 4
		yingliang=yingliang/2
		conn.execute "update �û� set ����=����+" &yinliang& ",����=����-500,����=����-200 where ����='"& aqjh_name &"'"
		conn.execute "update �û� set ����="& yinliang &" where ����='" & to1 & "'"
		fmrk=aqjh_name & "���"& to1 &"ƭ��������û�ɹ�����ȴƭ���������ֽ��һ��"& yinliang &","& aqjh_name &"�����½�500�㣬�����½�200�㡣"
	case 5
		conn.execute "update �û� set ����=����-50000000,����=����-500,����=����-500 where ����='"&aqjh_name&"'"
		fmrk=aqjh_name & "��թƭ"& to1 &"��˭֪"& to1 &"�͹ٸ���Ա��ʶ��"& aqjh_name &"�Ͽ����´����100�������Ӳ������ѣ����º�������˸��½�500�㡣"
	case 6
		conn.execute "update �û� set ����=����+500000,����=����-1000 where ����='"&aqjh_name&"'" 
		sql="insert into ����(����ID,����,��ò��) values ("& toid &",'" & to1 & "'," & meimao & ")"
set rs1=connt.execute(sql)
		fmrk=aqjh_name & "����Ƥ�����Ǻ��ã����ɴ�ƭ�ľͰ�"& to1 &"�����������ˣ��õ��ô���50�򣬵����½�1000��" 
	case 7
		conn.execute "update �û� set ����=0,����=����-500,����=����-30000,����=����-30000 where ����='"&aqjh_name&"'"
		fmrk=aqjh_name & "�ոս�"& to1 &"��ƭס��ȴ��"& to1 &"��һ����Ľ�߱���һ�٣���ʧ�����ֽ𣬵����½�500�������½�30000�������½�30000��"
     case 8
                yingliang=yingliang/2
		conn.execute "update �û� set ����=����+" &yinliang& ",����=����-500,����=����-200 where ����='"& aqjh_name &"'"
		conn.execute "update �û� set ����="& yinliang &" where ����='" & to1 & "'"
		fmrk=aqjh_name & "���"& to1 &"ƭ��������û�ɹ�����ȴƭ���������ֽ��һ��"& yinliang &","& aqjh_name &"�����½�500�㣬�����½�200�㡣"
     case 9
		conn.execute "update �û� set ״̬='��',��¼=now()+1/144,�¼�ԭ��='��ƭ�˿ڲ��ɹ���ץ' where ����='" & aqjh_name & "'"
		conn.execute "insert into l(b,c,d,a,e) values ('ϵͳ','" & to1 & "','����',now(),'" & fn1 & "')"
		fmrk=aqjh_name & "������ƭ"& to1 &"��ȴ���ٸ���Ա���֣�ץȥ�����ˡ�10���Ӻ�ſ��Ե�¼��"
		call boot(aqjh_name,"���Σ������ߣ�"&aqjh_name&","&fn1)
     case 10
                conn.execute "update �û� set ����=0,����=����-500,����=����-30000,����=����-30000 where ����='"&aqjh_name&"'"
		fmrk=aqjh_name & "�ոս�"& to1 &"��ƭס��ȴ��"& to1 &"��һ����Ľ�߱���һ�٣���ʧ�����ֽ𣬵����½�500�������½�30000�������½�30000��"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
set rs1=nothing
connt.close
set connt=nothing
end function
%>