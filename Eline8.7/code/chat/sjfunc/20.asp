<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_cf then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ҫ����ȼ�["&grade_cf&"]�ſ��Բ�����');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if Instr(Application("sjjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��" & towho & "�������������У����ܶ��䷢�ԣ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
	Response.end
end if
act=0
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
says="<font color=green>����⡿<font color=" & saycolor & ">"+cefen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���
function cefen(fn1,to1)
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert(��ʾ���������ɳ���10���ַ���');}</script>"
	Response.End
end if
if instr(fn1,"����")<>0 then
	Response.Write "<script language=JavaScript>{alert('������������ʲô��');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"����")<>0 or instr(fn1,"����")<>0 or instr(fn1,"����")<>0) then
	Response.Write "<script language=JavaScript>{alert('����:���ϡ�����������Ϊϵͳ�������벻Ҫʹ���������У�');}</script>"
	Response.End
end if
cefeng1=instr(says,"=")
cefeng2=instr(says,",")
cefeng3=instr(says,"grade")
cefeng4=instr(says,"���")
cefeng5=instr(says,"����")
cefeng6=instr(says,"�ٸ�")
if cefeng1<>0 or cefeng2<>0 or cefeng3<>0 or cefeng4<>0 or cefeng5<>0 or cefeng6<>0 then
	Response.Write "<script language=JavaScript>{alert('��������С������ң�����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ���� from �û� where ����='" & sjjh_name &"'" &" and ���='����'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����ʲôѽ����ɲ������ţ�');}</script>"
	Response.End
end if
mp=rs("����")
rs.close
rs.open "select * from �û� where ����='" & to1 &"'" &" and ����='" & mp & "'",conn,2,2
to1=rs("����")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]Ҳ�������ɵĵ���������ʲô��');}</script>"
	Response.End
end if
select case fn1
case "����"
	if rs("�ȼ�")<25 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������25�������ܷⳤ�ϣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As ���� from �û� where grade=4 and ���='����' and ����='"& mp &"'")
	musers=tmprs("����")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]�������ɵĳ�����8���ˣ���Ҫ�ٷ��ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ���='����',grade=4 where ����='" & to1 &"'"

case "����"
	if rs("�ȼ�")<20 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������20�������ܷ⻤����');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As ���� from �û� where grade=3 and ���='����' and ����='"& mp &"'")
	musers=tmprs("����")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=10 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]�������ɵĻ�����10���ˣ���Ҫ�ٷ��ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ���='����',grade=3 where ����='" & to1 &"'"

case "����"
	if rs("�ȼ�")<15 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������15�������ܷ�������');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As ���� from �û� where grade=2 and ���='����' and ����='"& mp &"'")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=16 then
		Response.Write "<script language=JavaScript>{alert('["&sjjh_name&"]�������ɵ�������16���ˣ���Ҫ�ٷ��ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ���='����',grade=2 where ����='" & to1 &"'"
case else
	conn.execute "update �û� set ���='"& fn1 &"',grade=1 where ����='" & to1 &"'"
end select
cefen=mp&"�����ţ�##���%%Ϊ" & mp & "��<font color=red><b>" & fn1 &"</b></font>�ɹ�,���ף�أ�"
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& sjjh_name &"','"& to1 &"','���','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>