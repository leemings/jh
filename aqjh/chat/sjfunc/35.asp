<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�������� xiulian
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(6)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]����û������¼�����������');}</script>"
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
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>������������<font color=" & saycolor & ">"+xiulian()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'�������� xiulian
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,��������,�ȼ�,����,����ʱ��,w5,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<>Application("aqjh_baowuname") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲢û�н�������["& Application("aqjh_baowuname") &"]��������');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û��������ɣ����["& s &"]���ٽ��в�����');}</script>"
	Response.End
end if
if rs("����")<2500  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������2500���޷�������');}</script>"
	Response.End
end if
zstemp=add(rs("w5"),"����",1)
if rs("��������")>=int(rs("�ȼ�")/Application("aqjh_baowuxl"))+1 then
	conn.execute "update �û� set ����='��',����=����+2000,������=������+��������,������=������+��������,�书��=�书��+��������,����=����+("& Application("aqjh_baowuyin") &"*��������),w5='"&zstemp&"' where ����='" & aqjh_name &"'"
	xll=rs("��������")
	conn.execute "update �û� set ��������=0,����ʱ��=now() where ����='" & aqjh_name &"'"
	xiulian="##<bgsound src=wav/xl.wav loop=1>ף����,������ĵĽ�������<font color=red>"& Application("aqjh_baowuname") &"</font>�������,�������޼�<font color=red>"& xll &"</font>��,���ֽ�+"& Application("aqjh_baowuyin")*xll &"��,����������2000��,����������ɣ��Զ���ʧ�ˡ���������<font color=blue>���ϵ��ӱ����б���<font color=red>1������</font>������~����</font>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<60 then
	s=60-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㻹û��������ɣ����["& s &"]���ٽ��в�����');}</script>"
	Response.End
end if
if rs("����")<2500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��������2500���޷�������');}</script>"
	Response.End
end if
randomize timer
r=int(rnd*7)+1
select case r
	case 1 '���
		n=int(rnd*1)+2
		conn.execute "update �û� set ���=���+"&n&" where ����='" & aqjh_name &"'"
		lw="���������ڷ�����������<font color=brown>���</font><font color=red>"&n&"</font>ö"
	case 2 '��
		n=int(rnd*1)+1
		conn.execute "update �û� set ��Ա��=��Ա��+"&n&" where ����='" & aqjh_name &"'"
		lw="��ɫ�ֻأ��쾪�ر䣬��������<font color=brown>��Ա��</font><font color=red>"&n&"</font>��"
	case 3 '�ֻ���
		n=int(rnd*1)+1
		temp=add(rs("w5"),"�ֻ���",n)
		conn.execute "update �û� set w5='"&temp&"' where ����='" & aqjh_name &"'"
		lw="�汦���汦��ֻ��һ�̹����������Ͻ���<font color=brown>�ֻ���</font><font color=red>"&n&"</font>��"
	case 4 '����
		n=int(rnd*50)+1
		conn.execute "update �û� set ����=����+"&n&" where ����='" & aqjh_name &"'"
		lw="���޳������޳����������齵�鸽���������������<font color=brown>����</font><font color=red>"&n&"</font>��"
	case 5 '֪��
		n=int(rnd*10)+1
		conn.execute "update �û� set ֪��=֪��+"&n&" where ����='" & aqjh_name &"'"
		lw="�Ķ����鶯���汦���⣬����<font color=brown>֪��</font><font color=red>"&n&"</font>��"
	case 6 '����
		n=int(rnd*1)+1
		temp=add(rs("w5"),"����",n)
		conn.execute "update �û� set w5='"&temp&"' where ����='" & aqjh_name &"'"
		lw="�������治�У�ֻ��һ�����������Ͻ���<font color=brown>����</font><font color=red>"&n&"</font>��"
end select 
conn.execute "update �û� set ��������=��������+1,����=����-2500,����ʱ��=now() where ����='" & aqjh_name &"'"
xiulian="##ӵ�н�������"& Application("aqjh_baowuname") &"�����������������:"& rs("��������") & "�ν��б�������....<font color=0088FF>�����汦�����亱�У�ÿ����һ�α�һ��Ʒ��</font><font color=blue>"&lw&"</font>!! "
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>