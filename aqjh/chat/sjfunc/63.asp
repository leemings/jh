<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
says="<font color=green>��������<font color=" & saycolor & ">"+huilai()+"</font></font>"
towho="���"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function huilai()
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))=0 then
	Response.Write "<script language=JavaScript>{alert('����û����Ϊ���룬�벻Ҫʹ�ô˹��ܣ�');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,��Ա�ȼ�,���,����,����ͷ��,�Ա�,��������,��ż,ͨ��,����,ְλ,ת��,���� from �û� where ����='" & aqjh_name &"'",conn,2,2
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
guojia=rs("����")
zhiwei=rs("ְλ")
jhjh=rs("����")
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
		aqjh_zm=aqjh_name&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&aqjh_jhdj&"|"&aqjh_id&"|"&hydj&"|0"&"|"&onlinexx(9)&"|"&mypeiou&"|"&myzs&"|"&guojia&"|"&zhiwei&"|"&jhjh
		onlinelist(i)=aqjh_zm
		exit for
	end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
aqjh_zanli=split(application("aqjh_zanli"),";")
for x=0 to UBound(aqjh_zanli)
dxname=split(aqjh_zanli(x),"|")
if dxname(0)="!"&aqjh_name&"!" then
	dxcz=aqjh_zanli(x)&";"
	application("aqjh_zanli")=replace(application("aqjh_zanli"),dxcz,"")
	 huilai=aqjh_name&"ȡ��������״̬���ָ�������״̬����"
	 exit for
end if
next
Application.UnLock
Response.Write "<script>parent.m.location.reload();parent.f2.document.af.addvalues.checked=false;<"&"/script>"
end function
%>