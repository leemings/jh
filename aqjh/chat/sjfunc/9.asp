<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_jg then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ����ȼ�["&grade_jg&"]�ſ��Բ�����');}</script>"
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
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����桿<font color=" & saycolor & ">"+jing(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function jing(fn1,to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('��TMD����������վ��Ҳ�Ҿ��棡');}</script>"
	Response.End
end if
fn1=trim(fn1)
jing="<font color=red><bgsound src=wav/warn.wav loop=1><b>[��������Ա]����ض�%%˵:" & fn1 & "! </b>(##)</font>"
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','����',now(),'" & fn1 & "')"
conn.close
set conn=nothing
end function
%>