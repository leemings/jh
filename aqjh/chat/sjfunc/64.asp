<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'����(��ϯ)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_jhdj<jhdj_xq then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ҫʹ�����鹦����Ҫ["&jhdj_xq&"]�����ϲſ���ʹ�ã�');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
says=replace(says,"\","")
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����顿<font color=" & saycolor & ">"+xinqing(mid(says,i+1))+"</font>"
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function xinqing(fn1)
fn1=trim(fn1)
if len(fn1)>3 then
	Response.Write "<script language=JavaScript>{alert('����״̬���Ȳ��ɴ���3���ַ���');}</script>"
	Response.End
end if
badword="վ��,����,����,�侫,��,��,ʺ,��,��,��,��,����,��,��,��,�,��,��,��,��,��,����,غ��,��,��Ƥ,��ͷ,��,�P,��,�H,��,��,��,����,����,��,����,����,��Ů,����,ʮ����,��,�Ҷ�,��,��ϯ,����,����,��־,������,�ٸ�"
bad=split(badword,",")
for i=0 to ubound(bad)-1
	if InStr(LCase(fn1),bad(i))<>0 then 
		Response.Write "<script language=JavaScript>{alert('��ʾ���������ֲ��ţ�');}</script>"
		Response.end
	end if
next
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,����,����ʱ��,��Ա�ȼ�,���,����,����ͷ��,�Ա�,��������,��ż,ͨ��,ת�� from �û� where ����='" & aqjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 and aqjh_grade<10 then
	s=3-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫʹ�����鹦�����["&s&"]���ٲ�����');}</script>"
	Response.End
end if
if rs("����")<1000 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ����1000���ú�������ɣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �û� set ����=����-1000,����ʱ��=now() where ����='" & aqjh_name &"'"
aqjh_id=rs("id")
hydj=rs("��Ա�ȼ�")
jhsf=rs("���")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="���"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="����"
end if
if rs("��ż")=Application("aqjh_user") and rs("�Ա�")="Ů" then
 jhsf="վ������"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="��������"
end if
if rs("ͨ��")=True then
	jhmp="ͨ����"
else
	jhmp=rs("����")
end if
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
myzs=rs("ת��")
mypeiou=rs("��ż")
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlinenum=UBound(onlinelist)
for i=1 to onlinenum step 1
	onlinexx=split(onlinelist(i),"|")
	if onlinexx(0)=aqjh_name then
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|"&onlinexx(8)&"|"&fn1&"|"&mypeiou&"|"&myzs
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application.UnLock
xinqing=aqjh_name&"�������ѵ�����״̬Ϊ["&fn1&"]����"
Response.Write ("<Script Language=JavaScript>parent.m.location.reload();</Script>")
end function
%>