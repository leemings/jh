<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'͵Ǯ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if aqjh_jhdj<=jhdj_tq then
	Response.Write "<script language=JavaScript>{alert('��ʾ��͵Ǯ��Ҫ["&jhdj_tq&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>����Ǯ��<font color=" & saycolor & ">"+touqian(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'͵Ǯ
function touqian(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,�ȼ�,����,��Ա�ȼ�,����,grade from �û� where ����='" & aqjh_name &"'" ,conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������˲����Բ���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ͵Ǯ!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,�ȼ�,����,��Ա�ȼ�,����,grade from �û� where ����='" & to1 &"'" ,conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ԶԳ����˲���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ǯ!');}</script>"
	Response.End
end if
if rs("�ȼ�")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���뽭��������Ǯ�𣿣�');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ͵����Ա��Ǯ����ѽ������');}</script>"
	Response.End
end if
to1=rs("����")
jhhy=rs("��Ա�ȼ�")
yin=rs("����")
rs.close
randomize timer
'��Ա2������͵Ǯ�ɹ�С��5%��
if jhhy<>0 then
	s=int(rnd*5)+1
else
	s=int(rnd*15)+1
end if
yin=int((yin/100))*s
r=int(rnd*7)
Select Case r
Case 1
	conn.execute "update �û� set ����=����-"&yin&" where ����='" & to1 &"'"
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>����һ�з���̽����,͵ȡ%%����"& s &"%,����:"& yin &"��!����С��ȫ����ʧ��!"
Case 2
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>����һ�з���̽����,����##��ٸ���ʶ,##��æ�����Ǹ,�����������!"
Case 3
	conn.execute "update �û� set ����=����-1500 where ����='" & aqjh_name &"'"
	touqian="##<bgsound src=wav/oh_no.wav loop=1>��͵ȡ%%������,��������ҷ�����,�����½�<font color=red>1500</font>��!<img src='img/daren.gif'>"
Case 4
	rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
	if rs("����")>50000 then
		conn.execute "update �û� set ����=����-50000 where ����='" & aqjh_name &"'"
		touqian="##<bgsound src=wav/qt.wav loop=1>͵%%��Ǯ������,ֻ�����´�㻨��<font color=red>5��</font>��,�ӹ��˽�,����ѽ!"
	else
		conn.execute "update �û� set ״̬='��' where ����='" & aqjh_name &"'"
		touqian="##<bgsound src=wav/oh_no.wav loop=1>�ոհ������%%��Ǯ��,�ͱ��ٸ����˷�����,������׽ȥ������"
		call boot(aqjh_name,"͵Ǯ�������ߣ�"&to1&"��������͵��")
	end if
	rs.close
case else
	conn.execute "update �û� set ����=����-"&yin&" where ����='" & to1 &"'"
	conn.execute "update �û� set ����=����+"&yin&" where ����='" & aqjh_name &"'"
	touqian="##<bgsound src=wav/xiaotou.wav loop=1>����һ�з���̽����,͵ȡ%%����"& s &"%,����:"& yin &"��!%%���,��Ҫ����Ӵ!"
End Select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
