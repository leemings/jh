<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ϲ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
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
says="<font color=green>����ϲ�ǡ�<font color=" & saycolor & ">"+fen(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ϲ��
function fen(fn1,to1)
fn1=int(abs(fn1))
if fn1<500 or fn1>1000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��500-1000��');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,��ż,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<50000000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô���Ǯ�ָ������?');}</script>"
	Response.End
end if
if rs("��ż")="��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹û���,��ʲôϲ�ǰ���');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���"& s &"���ٲ���,�ɱ����ţ�');</script>"
	response.end
end if

Response.Buffer=true
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
for i=0 to x
conn.Execute "update �û� set ����=����+" & fn1 & " where ����='" & online(i) & "'"
next
conn.Execute "update �û� set ����=����-" & fn1*1000 & ",����=����+1000,����ʱ��=now() where ����='" & aqjh_name & "'"
fen="##<bgsound src=wav/FAQIAN.wav loop=1>�����["&rs("��ż")&"]�������ϲ������,�ó�"&fn1*1000&"������ҷ�ϲ��!ÿ�˷ֵ�����"&fn1&",��������1000,��ҹ�ϲ��(��)����!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
