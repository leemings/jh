<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ʹ����ҩ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
'#####���䴦��#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
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
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
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
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ����ҩ��');}</script>"
	Response.End
end if
f=Minute(time())
if f<00 or f>10 then
	Response.Write "<script language=JavaScript>{alert('���ڲ���ʹ����ҩʱ�䣬ʹ����ҩΪ1-10���ӣ�');window.close();}</script>"
	Response.End 
end if
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr(";ըҩ;��ʬˮ;��Ѫ��;�Ƴ�;����ɢ;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�������Լ����в�����');</script>"
	Response.End
end if
if to1="���" and instr("ըҩ;��ʬˮ;��Ѫ��;�Ƴ�;����ɢ;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('��"&fn1&"�����ܴ�ҽ��в�����');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade from �û� where ����='" & to1 &"'",conn,2,2
if rs("grade")>=6 and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
rs.close
rs.open "select grade from �û� where  ����='" & aqjh_name &"'",conn,2,2
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if
rs.close
fn1=trim(fn1)
rs.open "select w8 from �û� where ����='"&aqjh_name&"'",conn,2,2
mycard=abate(rs("w8"),fn1,1)
rs.close
rs.open "select d,e from x where  a='"&fn1&"'",conn,2,2
cz=rs("d")
jjsj=rs("e")
rs.close
select case fn1
case "������"
	conn.Execute ("update �û� set ɱ����=0 where ����='" & aqjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ����ҩ[<b>"&fn1&"</b>]ɱ������0...</font>"
case "ըҩ","��ʬˮ"
	conn.Execute ("update �û� set "&cz&"=int("&cz&"/"&jjsj&") where ����='" & to1 &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ��[<b>"&fn1&"</b>]����%%��ʹ��%%��"&cz&"�½�1/"&jjsj&"��...</font>"
case "��Ѫ��","�Ƴ�","����ɢ"
	rs.open "select "&cz&" from �û� where  ����='"&to1&"'",conn,2,2
	abc=rs(cz)
	rs.close
	conn.Execute ("update �û� set "&cz&"=int("&cz&"/"&jjsj&") where ����='" & to1 &"'")
	conn.Execute ("update �û� set "&cz&"="&cz&"+int("&abc/jjsj&") where ����='" & aqjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ��[<b>"&fn1&"</b>]����%%��ʹ��%%��"&cz&"�½�1/"&jjsj&"��,��ȡ�Է�("&cz&")��:+<font color=red>"& int(abc/jjsj) &"</font>�㡭��</font>"
case else
	conn.Execute ("update �û� set "&cz&"="&cz&"+"&jjsj&" where ����='" & aqjh_name &"'")
	peiyao="<font color=green>��ʹ����ҩ��<font color=" & saycolor & ">##ʹ����ҩ[<b>"&fn1&"</b>]("&cz&")���ӣ�<font color=red>+"&jjsj&"</font>��...</font>"
end select
'ɾ��ҩƷ
conn.execute "update �û� set w8='"&mycard&"' where  ����='"&aqjh_name&"'"
set rs=nothing	
conn.close
set conn=nothing
end function
%>