<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_nj then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ҫ["&jhdj_nj&"]���ſ��Բ�����');}</script>"
	Response.End
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
mycz=trim(left(says,i-1))
if mycz="/�" then
	says="<font color=green>�����<font color=" & saycolor & ">"+nianjing(towho)+"</font>"
else
	says="<font color=green>�����崺�硿<font color=" & saycolor & ">"+chunfeng(towho)+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�
function nianjing(to1)
if to1="���" or to1=Application("aqjh_automanname") then to1=aqjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if to1<>aqjh_name then
	rs.open "select ����,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
	if rs("����")="����" or rs("ͨ��")=True then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���Է�Ϊ����ɱ�֣���ͨ���������Բ���!",1)
	end if
	if rs("����")>10000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���Է����´���1�򲻿ɲ���!",1)
	end if
	rs.close
end if
rs.open "select ����,����,����,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<10000 or rs("����")<1000 or rs("����")<>"����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���������˲ſ��Բ���!\n��Ҫ1������1000����!",1)
end if
conn.execute "update �û� set ����=����-10000,����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����+150 where ����='" & to1 &"'"
if to1<>aqjh_name then
	nianjing="������##ϯ�ض�����˫�ֺ�ʵ�����������дʣ��ϡ��֡�٢���𡭡�%%��������$$redb150$$b�㣬������##��������$$blueb1000$$b�㣡"
else
	nianjing="������##���ڷ�����ǰ������������˵�Լ��ں쳾�е���񡭡����ϵ��İѷ��������ӵ���$$redb150$$b�㡭��"
end if	
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

'���崺��
function chunfeng(to1)
to1=trim(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("����")="����" or rs("ͨ��")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���Է�Ϊ����ɱ�֣���ͨ���������Բ���!",1)
end if
if rs("����")>10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���Է���������1�򲻿ɲ���!",1)
end if
rs.close
rs.open "select ����,�Ա�,����,����,����,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("�Ա�")="��" then
	if rs("����")<10000 or rs("����")<1000 or rs("����")<>"����" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���������˲ſ��Բ���!\n�г�������Ҫ1������1000����!",1)
	end if
	conn.execute "update �û� set ����=����-10000,����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
	chunfeng="�г�����##����һ������(%%)���������֮�֣�Ϊ�����ˡ��������ϼ�������%%�����ָ�������$$redb2000$$b�㣬������##��������$$blueb1000$$b�㣡"
else
	if rs("����")<1000 or rs("����")<1000 or rs("����")<>"����" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���������˲ſ��Բ���!\nŮ��������Ҫ1������1000����!",1)
	end if
	conn.execute "update �û� set ����=����-10000,����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
	chunfeng="Ů������##����һ������(%%)���������֮�֣���������֭Ϊ�����ˣ�%%�����ָ�������$$redb1500$$b�㣬������##��������$$blueb1000$$b�㣡"
end if
conn.execute "update �û� set ����=����+2000 where ����='" & to1 &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
