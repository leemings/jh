<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��Ѩ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_dx then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ѩ��Ҫ����ȼ�["&grade_dx&"]�ſ��Բ�����');}</script>"
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>����Ѩ��</font><font color=" & saycolor & ">"+dian(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function dian(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade,���� from �û� where ����='" & to1 &"'",conn
to1=rs("����")
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��û������ˣ����ǲ��ǿ����ˣ�');}</script>"
	Response.End
end if
if rs("grade")<aqjh_grade or aqjh_name=Application("aqjh_user") then
	conn.execute "update �û� set ��¼=now()+0.02,״̬='��Ѩ',�¼�ԭ��='"&aqjh_name&" ��Ѩ:"&fn1&"' where ����='" & to1 &"'"
	dian="##��%%ʹ���˵�Ѩ����%%�����ز����ˡ�����28���Ӳſ������ߡ���ԭ��"&fn1
	call boot(to1,"��Ѩ�������ߣ�"&aqjh_name&","&fn1)
else
	dian="##��%%ʹ���˵�Ѩ�����������ǹٸ��и߼�����Ա,�����û�а취��"
end if
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','��Ѩ','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
