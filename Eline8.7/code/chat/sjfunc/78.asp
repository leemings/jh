<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_nj then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ҫ["&jhdj_nj&"]���ſ��Բ�����');}</script>"
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
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
mycz=trim(left(says,i-1))
if mycz="/����" then
	says="<font color=green>�����<font color=" & saycolor & ">"+nianjing(towho)+"</font>"
else
	says="<font color=green>�����崺�硿<font color=" & saycolor & ">"+chunfeng(towho)+"</font>"
end if
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�
function nianjing(to1)
if to1="���" or to1=Application("sjjh_automanname") then to1=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if to1<>sjjh_name then
	rs.open "select ����,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
	if rs("����")="����" or rs("ͨ��")=True then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���Է�Ϊ����ɱ�֣���ͨ���������Բ�����",1)
	end if
	if rs("����")>10000 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���Է����´���1�򲻿ɲ�����",1)
	end if
	rs.close
end if
rs.open "select ����,����,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*20)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*20)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<500000 or rs("����")<1000 or rs("����")<>"����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���������˲ſ��Բ���!\n��Ҫ50������1000����!",1)
end if
conn.execute "update �û� set ����=����-500000,����=����-1000,����ʱ��=now() where ����='" & sjjh_name &"'"
conn.execute "update �û� set ����=����+150 where ����='" & to1 &"'"
if to1<>sjjh_name then
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
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("����")="����" or rs("ͨ��")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���Է�Ϊ����ɱ�֣���ͨ���������Բ�����",1)
end if
if rs("����")>10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call mess("��ʾ���Է���������1�򲻿ɲ�����",1)
end if
rs.close
rs.open "select ����,�Ա�,����,����,����,����ʱ�� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
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
	if rs("����")<500000 or rs("����")<1000 or rs("����")<>"����" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���������˲ſ��Բ�����\n�г�������Ҫ50������1000������",1)
	end if
	conn.execute "update �û� set ����=����-500000,����=����-1000,����ʱ��=now() where ����='" & sjjh_name &"'"
	chunfeng="�г�����##����һ������(%%)���������֮�֣�Ϊ�����ˡ��������ϼ�������%%�����ָ�������$$redb2000$$b�㣬������##��������$$blueb1000$$b�㣡"
else
	if rs("����")<100000 or rs("����")<1000 or rs("����")<>"����" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		call mess("��ʾ���������˲ſ��Բ�����\nŮ��������Ҫ10������1000������",1)
	end if
	conn.execute "update �û� set ����=����-100000,����=����-1000,����ʱ��=now() where ����='" & sjjh_name &"'"
	chunfeng="Ů������##����һ������(%%)���������֮�֣���������֭Ϊ�����ˣ�%%�����ָ�������$$redb1500$$b�㣬������##��������$$blueb1000$$b�㣡"
end if
conn.execute "update �û� set ����=����+2000 where ����='" & to1 &"'"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
