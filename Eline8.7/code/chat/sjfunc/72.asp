<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="chatconfig.asp"-->
<%'ʹ����ҩ��wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####���䴦��#####
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ����ҩ��');}</script>"
	Response.End
end if
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="����E��" then
	Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ�����в�����ʹ����ҩ��');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=peiyao(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ʹ����ҩ
function peiyao(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if sjjh_name=to1 and instr(";ըҩ;��ʬˮ;��Ѫ��;�Ƴ�;����ɢ;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�������Լ����в�����');</script>"
	Response.End
end if
if to1="���" and instr("ըҩ;��ʬˮ;��Ѫ��;�Ƴ�;����ɢ;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�����ܴ�ҽ��в�����');</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
fn1=trim(fn1)
rs.open "select ��Ա��,w8 from �û� where  ����='"&sjjh_name&"'",conn,2,2
mycard=abate(rs("w8"),fn1,1)
rs.close
rs.open "select d,e from x where  a='"&fn1&"'",conn,2,2
cz=rs("d")
jjsj=rs("e")
rs.close
select case fn1
case "������"
	conn.Execute ("update �û� set ɱ����=0 where ����='" & sjjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ����ҩ[<b>"&fn1&"</b>]ɱ������0...</font>"
case "ըҩ","��ʬˮ"
	conn.Execute ("update �û� set "&cz&"=int("&cz&"/"&jjsj&") where ����='" & to1 &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ��[<b>"&fn1&"</b>]����%%��ʹ��%%��"&cz&"�½�1/"&jjsj&"��...</font>"
case "��Ѫ��","�Ƴ�","����ɢ"
	rs.open "select "&cz&" from �û� where  ����='"&to1&"'",conn,2,2
	abc=rs(cz)
	rs.close
	conn.Execute ("update �û� set "&cz&"=int("&cz&"/"&jjsj&") where ����='" & to1 &"'")
	conn.Execute ("update �û� set "&cz&"="&cz&"+int("&abc/jjsj&") where ����='" & sjjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ��[<b>"&fn1&"</b>]����%%��ʹ��%%��"&cz&"�½�1/"&jjsj&"��,��ȡ�Է�("&cz&")��:+<font color=red>"& int(abc/jjsj) &"</font>�㡭��</font>"
case else
	conn.Execute ("update �û� set "&cz&"="&cz&"+"&jjsj&" where ����='" & sjjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ����ҩ[<b>"&fn1&"</b>]("&cz&")���ӣ�<font color=red>+"&jjsj&"</font>��...</font>"
end select

'ɾ���Լ���Ƭ����¼
conn.execute "update �û� set w8='"&mycard&"' where  ����='"&sjjh_name&"'"
set rs=nothing	
conn.close
set conn=nothing
end function
%>
