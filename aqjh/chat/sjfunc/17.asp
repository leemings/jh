<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_zs then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ["&jhdj_zs&"]���ſ��Բ�����');}</script>"
	Response.End
end if
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"�������" then
 Response.Write "<script language=javascript>{alert('��ʾ��������Ʒ��ȥ������');}</script>"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����͡�<font color=" & saycolor & ">"+zen(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function zen(fn1,to1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
zt=split(fn1,"|")
if ubound(zt)<>2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ�����|��Ʒ��|���� ');}</script>"
	Response.End 
end if
if not isnumeric(zt(2)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ʹ�����֣�');}</script>"
	Response.End 
end if
lb=trim(zt(0))
zswupin=trim(zt(1))
wusl=abs(int(clng(zt(2))))
if wusl=0 or wusl>100 and instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��100��');}</script>"
	Response.End 
end if
if lb<>"w1" and lb<>"w2" and lb<>"w3" and lb<>"w4" and lb<>"w5" and lb<>"w6" and lb<>"w7" and lb<>"w8" and lb<>"w10" then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ�����ȷ��');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,"&lb&" from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("����")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ʒ��Ҫ5000�������ѣ�');}</script>"
	Response.End
end if
zstemp=abate(rs(lb),zswupin,wusl)
conn.execute "update �û� set "&lb&"='"&zstemp&"',����=����-5000 where  ����='"&aqjh_name&"'"
rs.close
rs.open "select "&lb&" from �û� where ����='"&to1&"'",conn,2,2
zstemp=add(rs(lb),zswupin,wusl)
conn.execute "update �û� set "&lb&"='"&zstemp&"' where  ����='"&to1&"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','����',now(),'����["&zswupin&"]"&wusl&"��')"
zen="##���Լ���[<b><font color=blue>" & zswupin & "</font></b>]���͸���%%["&wusl&"]����%%���Ǹ�л,�˴β����ɹ�!��"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
