<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->

<%'Ѱ�ҽ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_xunshuijing then
	Response.Write "<script language=JavaScript>{alert('Ѱ�ҽ�ң���Ҫ�����ȼ�["&jhdj_xunshuijing&"]���Ĳſ��Բ�����');}</script>"
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>��Ѱ�ҽ�ҡ�<font color=" & saycolor & ">"+xunshuijing(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'Ѱ�ҽ��
function xunshuijing(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����ʱ��,֪��,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<(int(rnd*5)+1) then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=(int(rnd*5)+1)-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
fla=rs("����")
if rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ķ��������޷�ʩչѽ������Ҳ��5000�㰡��');window.close();}</script>"
	response.end
end if
if rs("֪��")<20 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����֪�ʲ����޷�ʩչѽ������Ҳ��20�㰡��');window.close();}</script>"
	response.end
end if
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
dj=rs("�ȼ�")
if dj<30 then
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ȼ�Ϊ30���ſ��Խ���Ѱ�ҽ�ң�');}</script>"
	Response.End
end if
rs.close
conn.execute "update �û� set ����ʱ��=now(),����=����-5000,֪��=֪��-20 where ����='"&aqjh_name&"'"
randomize 
r=int(rnd*6)+1
select case r
case 1
xunshuijing=aqjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲô���," & aqjh_name & "���1000�㷨��!" 
	conn.execute "update �û� set ����=����-1000 where ����='" & aqjh_name &"'"
	
case 2
xunshuijing=aqjh_name & "�����ˣ���Ѱ���˽������������ҵ���Ѱ����100�㷨��" & aqjh_name & "�����˷���100��!" 
	conn.execute "update �û� set ����=����+100 where ����='" & aqjh_name &"'"

case 3
xunshuijing=aqjh_name & "��Ѱ���˽����������������ҵ���[1]�����," & aqjh_name & "���1000�㷨��!" 
	conn.execute "update �û� set ����=����-1000,���=���+1 where ����='" & aqjh_name &"'"

case 4
xunshuijing=aqjh_name & "�����ˣ���Ѱ���˽������������ҵ���Ѱ�����100�㷨��" & aqjh_name & "�����˷���100��!" 
	conn.execute "update �û� set ����=����+100 where ����='" & aqjh_name &"'"

case 5
xunshuijing=aqjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲô���," & aqjh_name & "���1000�㷨��!" 
	conn.execute "update �û� set ����=����-1000 where ����='" & aqjh_name &"'"

case 6
xunshuijing=aqjh_name & "�ÿ�ϧ����Ѱ���˽�����������Ҳû���ҵ�ʲô���," & aqjh_name & "���1000�㷨��!" 
	conn.execute "update �û� set ����=����-1000 where ����='" & aqjh_name &"'"
	rs.close
	
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>