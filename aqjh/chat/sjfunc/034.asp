<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�Է�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_jhdj<jhdj_qk then
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ҫ["&jhdj_qk&"]���ſ��Բ�����');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('��ʾ���������������˻��Լ�ѽ��');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=blue>����Ǯ��͡�<font color=green>"+qw(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function qw(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ��,����,����,�Ա� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<10 or rs("����")<50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����²���10����������50����������Է���');}</script>"
	Response.End
end if
sj=DateDiff("s",rs("����ʱ��"),now())
'if sj<20 then
'	s=20-sj
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('���˼�Ҫ20��һ�Σ����"& s &"������,�ɱ����ţ�');</script>"
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
if rs("����")<800 or rs("����")<500 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ҫ�ķ�800������500������"& to1 &"̫���ˣ���');</script>"
	response.end
end if
conn.execute "update �û� set ����ʱ��=now() where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("aqjh_qw")=aqjh_name&"|"&to1&"|"&now()
Application.UnLock
randomize()
regjm=int(rnd*3144998)
aa="<input class=btn style='font-size: 9pt; background-color: #FFCCCC; border-style: ridge' type=button value="
cc="onclick=javascript:;ai"&regjm&".disabled=1;jj"&regjm&".disabled=1;window.open('hyqk.asp?zl="
qw="[##]����һ���Ǯ��{%%}˵��"& fn1 & aa &"'�Է�' "& cc &"�Է�','d') name=ai"&regjm&">"& aa &"'�����' "& cc &"�����','d') name=ai"&regjm&">"& aa &"'����' "& cc &"����','d') name=ai"&regjm&">"& aa &"'�ܾ�' "& cc &"�ܾ�','d') name=jj"&regjm&">��Щ������һ�ְɣ�������~��"
end function
%>