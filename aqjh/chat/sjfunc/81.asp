<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'Ǭ��һ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
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
rs.open "select ����,����ʱ��,grade,����,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=2-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲��ܲ�������K�����˵��');}</script>"
	Response.End
end if

if rs("����")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô���Ǯ�𣿿�����˵��');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ���Ա!');}</script>"
	Response.End
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ڱ�����!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,grade,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("����")="����" and aqjh_grade<10 and rs("ͨ��")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɹ���������!');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɹ����ٸ���Ա!');}</script>"
	Response.End
end if
if rs("����")=True and aqjh_grade<10 and rs("ͨ��")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ڱ�����!');}</script>"
	Response.End
end if
if rs("����")<=1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&to1&"]�������������ˣ��㲻Ҫ�ٹ�����!');}</script>"
	Response.End
end if
randomize ()
m1 = Int(100 * Rnd)+100
gjtl=int(fn1/m1)
gjtl=rs("����")-gjtl
if gjtl<1000 then 	gjtl=1000
conn.execute "update �û� set ����="&gjtl&" where ����='" & to1 &"'"
conn.execute "update �û� set ����=����-" & fn1 & ",����ʱ��=now() where ����='" & aqjh_name &"'"
moneygj="##�ӿڴ����ó�����" & fn1 & "����ʹ����һ���������Ǭ��һ��,���е�������һ�����ε����������%%��ͷת��%%����ֻʣ��[$$redb"&gjtl&"$$b]��~~"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
