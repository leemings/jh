<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'���ܽӴ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_qr then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ܽӴ���Ҫ["&jhdj_qr&"]���ſ��Բ�����');}</script>"
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
if towho=aqjh_name or towho=application("aqjh_automanname") then
	towho="���"
else
	call dianzan(towho)
end if
if towho="���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ������ɣ����׽������е���Ů����������Ź֣�');}</script>"
	Response.End
end if
if application("aqjh_tw")<>"" then
	wddata=split(application("aqjh_tw"),"|")
	if ubound(wddata)=2 then
		sj=wddata(2)
	end if
	erase wddata
	nowsj=DateDiff("s",sj,now())
	if nowsj<10 then
		jgsj=10-nowsj
		Response.Write "<script language=JavaScript>{alert('��һ�����ܽӴ���δ���������ٵ�"& jgsj &"�룡');}</script>"
		Response.End
	end if
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","��")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����ܽӴ���<font color=" & saycolor & ">"+tw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ǩ
function tw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ��,����,����,����,����,�Ա� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<500 or rs("����")<500 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����²���500����������500��������ɫ�ˣ�');}</script>"
	Response.End
end if
if rs("����")<800 or rs("����")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ܽӴ�Ҫ�ķ�800������500��������̫���˰ɣ���');</script>"
	response.end
end if
sj=DateDiff("s",rs("����ʱ��"),now())
'if sj<100 then
'	s=100-sj
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('���ܽӴ�Ҫ100��һ�Σ����"& s &"��,�ɱ����ţ�');</script>"
'	response.end
'end if
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
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ܽӴ�ֻ������������!');}</script>"
	Response.End
end if
if rs("����")<1 or rs("����")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����²���100����������100��ɫ��������Ҳ̫û�����ˣ�');}</script>"
	Response.End
end if
if rs("����")<800 or rs("����")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����ܽӴ�Ҫ�ķ�800������500������"& to1 &"̫���ˣ���');</script>"
	response.end
end if
conn.execute "update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_tw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
randomize()
regjm=int(rnd*3144998)
aa="<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value="
cc="onclick=javascript:;jgj"&regjm&".disabled=1;qjd"&regjm&".disabled=1;jhdt"&regjm&".disabled=1;jj"&regjm&".disabled=1;window.open('wd1.asp?zl="
tw="[##]��{%%}<img src='img/xu1.gif'>��������ܽӴ������󣺱�������Ӵ���������֤���֡���˵����ô���㣡������һ��ֵ�ð���"& fn1 & aa &"'������' "& cc &"������','d') name=jgj"&regjm&">"& aa &"'ȥ�Ƶ�' "& cc &"ȥ�Ƶ�','d') name=qjd"&regjm&">"& aa &"'�ڽ�������' "& cc &"�ڽ�������','d') name=jhdt"&regjm&">"& aa &"'�ܾ�' "& cc &"�ܾ�','d') name=jj"&regjm&">"
end function
%>