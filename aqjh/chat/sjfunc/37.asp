<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��Ѩ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_grade<4 then
	Response.Write "<script language=JavaScript>{alert('����ʲôѽ����Ĺ���ȼ��ɲ���ѽ��');}</script>"
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
says="<font color=green>����Ѩ��<font color=" & saycolor & ">"+ya(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��Ѩ
function ya(to1,fn1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('���棺�㲻���Զ�վ��������');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
grade=rs("grade")
menpai=rs("����")
denji=rs("grade")
rs.close
rs.open "select ����,grade from �û� where ����='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��û������ˣ����ǲ��ǿ����ˣ�');}</script>"
	Response.End
end if
if denji<6 and menpai<>rs("����") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������["&to1&"]Ҳ�����������ɵ�ѽ��');}</script>"
	Response.End
end if
if rs("grade")<grade or aqjh_name=Application("aqjh_user") then
	Application.Lock
	application("aqjh_dianxuename")=application("aqjh_dianxuename")&to1&"|"&now()&"|"&";"
	Application.UnLock
	ya="##��%%ʹ������Ѩ����" & to1 & "�����ز����ˡ���"
else
	ya="##��%%ʹ������Ѩ����������ĵȼ������˼ң�û�취��"
end if
'��¼
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','��Ѩ','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>