<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�ͽ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_jhdj<jhdj_sq then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ͽ����Ҫ["&jhdj_sq&"]���ſ��Բ�����');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(0)<>"�������" then
 Response.Write "<script language=javascript>{alert('��ʾ���ͽ����ȥ����������û���ڣ�');}</script>"
 Response.End
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
says="<bgsound src=wav/THANKS.wav loop=1><font color=green>���ͽ�ҡ�<font color=" & saycolor & ">"+givegold(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�ͽ��
function givegold(fn1,to1)
fn1=int(abs(fn1))
if (fn1<1 or fn1>100) and aqjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ͽ������1�������100����');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���,����,�Ṧ,grade FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("���")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô��Ľ���𣿿�����˵��');}</script>"
	Response.End
end if
if rs("����")<1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ͽ����Ҫ1000�����ӵĳ��طѣ�');}</script>"
	Response.End
end if
if rs("grade")<4 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ�����Ÿ�����Ա�ſ����ͽ�ң�');}</script>"
	Response.End
end if
if rs("�Ṧ")<2000 then
Response.Write "<script language=JavaScript>{alert('��ʾ�����ͽ����Ҫ2000���Ṧ�������ѣ���');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set �Ṧ=�Ṧ-2000,���=���-" & fn1 & ",����=����-1000 where ����='" & aqjh_name &"'"
conn.execute "update �û� set ���=���+" & fn1 & " where ����='" & to1 &"'"
givegold="##��" & fn1 & "����͸���%%[$$redb"&fn1&"$$b]�����˴����Ͳ����ɹ�!"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','����','�ͽ��"&fn1&"��')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
onn.close
set conn=nothing
end function
%>