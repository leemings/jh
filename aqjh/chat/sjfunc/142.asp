<!--#include file="sjfunc.asp"-->
<%'š��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if towho="���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ������ɣ������˭ʹ�ð���');}</script>"
	Response.End
end if
act=1
towhoway=0
says="<font color=green>���ҷ����Ρ�<font color=" & saycolor & ">"+tw(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tw(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ա�,��ż,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ĵ���̫�ܻ��ˣ����±�������');}</script>"
	Response.End
end if
mysex=rs("�Ա�")
mywf=rs("��ż")
if mysex="��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ֹ����Ƕ���ʹ�õİ���');}</script>"
	Response.End
end if
if mywf="��" then
 tw="##:�㻹û�н���أ�����ʹ�üҷ�����Ŷ^_^"
 exit function
end if
if mywf<>to1 then
 tw="##:�㲻�ܶ�%%ʹ�üҷ����Σ�%%�ɲ������Ϲ�Ŷ^_^"
 exit function
end if
rs.close
rs.open "select ����,���� FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ǽ������ˣ������Զ���ʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ϲ��������Ѿ��ܵ��˾ͷŹ���һ�ΰɣ�');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ϲ��������Ѿ������˾ͷŹ���һ�ΰɣ�');</script>"
	response.end
end if
conn.execute "update �û� set ����=����-30,����ʱ��=now() where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����-50,����=����-40 where ����='"&to1&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
tw="����##���ߺߵض��Ϲ�%%ʹ���˼ҷ�<img src=pic/zm1.gif>�����㻹�ҳ�ȥ��˵�˵�������Ϲ���ð���ǣ����������������½�40��������ʧ50�����ڷ������᲻�٣������½�30"
end function
%>