<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���������wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_jhdj<jhdj_tw then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������Ҫ["&jhdj_chi&"]���ſ��Բ�����');}</script>"
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
if towho=sjjh_name or towho=application("sjjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if towho="���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������������飬��������Ź֣�');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����衿<font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ǩ
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����ʱ��,����,����,����,����,�Ա� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")<1000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����²���1000����������5000��������������飡');}</script>"
	Response.End
end if
if rs("����")<8000 or rs("����")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������Ҫ�ķ�8000������5000��������̫���˰ɣ���');</script>"
	response.end
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<100 then
s=100-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��������Ҫ300��һ�Σ����"& s &"��,�ɱ����ţ�');</script>"
response.end
end if
mysex=rs("�Ա�")
rs.close
rs.open "select ����,����,����,����,�Ա� FROM �û� WHERE ����='" & to1 &"'" ,conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ǽ������ˣ������Զ���ʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
if rs("�Ա�")=mysex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������ֻ�����������ѣ�');}</script>"
	Response.End
end if
if rs("����")<1000 or rs("����")<5000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����²���1000����������5000���������Ҳ̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<8000 or rs("����")<5000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������Ҫ�ķ�8000������5000������"& to1 &"̫���ˣ���');</script>"
	response.end
end if

conn.execute "update �û� set ����ʱ��=now() where ����='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("sjjh_tw")=sjjh_name&"|"&to1
Application.UnLock
randomize()
regjm=int(rnd*3144998)
tw="<bgsound src=wav/ring.WAV loop=1>[##]��{%%}<img src='img/yu.gif' width='70' height='55'>����������������Ӵ����������Ʊ���������������˵úܣ�<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='������' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=������','d') name=yuncheng"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='̽��' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=̽��','d') name=yinyuan"&regjm&"><input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value='��˹��' onClick=javascript:;yuncheng"&regjm&".disabled=1;yinyuan"&regjm&".disabled=1;shiye"&regjm&".disabled=1;window.open('wd.asp?fromname=" & sjjh_name &"&toname="&to1&"&qq=��˹��','d') name=shiye"&regjm&">"
end function
%>