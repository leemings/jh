<%@ LANGUAGE=VBScript codepage ="936" %><!--#include file="sjfunc.asp"-->
<%'���Ҳ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_cf then
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
if Instr(Application("aqjh_useronlinename"&nowinroom)," "&towho&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��" & towho & "�������������У����ܶ��䷢�ԡ�');parent.f2.document.af.towho.value='���';parent.f2.document.af.towho.text='���';parent.m.location.reload();</script>"
	Response.end
end if
act=0
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
says="<font color=green>�����Ҳ�⡿<font color=" & saycolor & ">"+cefen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���
function cefen(fn1,to1)
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert(��ʾ���������ɳ���10���ַ���');}</script>"
	Response.End
end if
if instr(fn1,"����")<>0 then
	Response.Write "<script language=JavaScript>{alert('�������ʵۣ�����ʲô��');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"ة��")<>0 or instr(fn1,"����")<>0 or instr(fn1,"����")<>0) then
	Response.Write "<script language=JavaScript>{alert('����:ة�ࡢ����������Ϊϵͳ�������벻Ҫʹ���������У�');}</script>"
	Response.End
end if




Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� from �û� where ����='" & aqjh_name &"'" &" and ְλ='����'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����ʲôѽ����ɲ��ǻʵۣ�');}</script>"
	Response.End
end if
guojia=rs("����")
rs.close
rs.open "select * from �û� where ����='" & to1 &"'" &" and ����='" & guojia & "'",conn,2,2
to1=rs("����")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]�������������������ʲô��');}</script>"
	Response.End
end if
select case fn1
case "ة��"
	if rs("�ȼ�")<100 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������100�������ܷ�ة��!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As ���� from �û� where ְλ='ة��' and ����='"& guojia &"'")
	musers=tmprs("����")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=4 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]������Ĺ����Ѿ���2��ة���ˣ���Ҫ�ٷ��ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ְλ='����'where ����='" & to1 &"'"

case "����"
	if rs("�ȼ�")<80 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������80�������ܷ�λ����Ľ���!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As ���� from �û� where ְλ='����' and ����='"& guojia &"'")
	musers=tmprs("����")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]������Ĺ����Ѿ���4�������ˣ���Ҫ�ٷ��ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ���='����' where ����='" & to1 &"'"

case ""
	if rs("�ȼ�")<60 then
		Response.Write "<script language=JavaScript>{alert('["&to1&"]������60�����������������ܱ�����Ĺ���ô!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As ���� from �û� where ְλ='����' and ����='"& guojia &"'")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=16 then
		Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]������Ĺ����Ѿ���8���������㹻ǿ���ˣ�');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �û� set ְλ='����' where ����='" & to1 &"'"
case else
	conn.execute "update �û� set ְλ='"& fn1 &"' where ����='" & to1 &"'"
end select
cefen=guojia&"�Ļʵ�����گ�飺������ˣ��ʵ�گԻ��##���%%Ϊ" & guojia & "��<font color=red><b>" & fn1 &"</b></font>�ɹ�, %%�Ӵ˹��˺�ͨ����������"
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','���','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>