<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'ʹ�á�wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
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
call dianzan(towho)
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
says="<font color=green>��ʹ��ҩƷ��<font color=" & saycolor & ">"+eat(mid(says,i+1))+"</font>"
towhoway=1
towho=sjjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ʹ��
function eat(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('��ʾ�����ɣ�������ʲô���뵷���𣿣�');}</script>"
Response.End 
end if
zt=split(fn1,"&")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ����Ʒ��|���� ');}</script>"
	Response.End 
end if
if not isnumeric(zt(1)) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ����������������ʹ�����֣�');}</script>"
	Response.End 
end if
'���䴦��
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="����E��" and Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<=22) then
	if instr(fn1,"ǧ���˲�")=0 or instr(fn1,"������֥")=0 then
		Response.Write "<script language=JavaScript>{alert('��ʾ���ᱦ�ڼ�ֻ���Է���[ǧ���˲�]��[������֥]����ҩƷ����������������');}</script>"
		Response.End
	end if
else
	if instr(fn1,"ǧ���˲�")<>0 or instr(fn1,"������֥")<>0 then
		Response.Write "<script language=JavaScript>{alert('��ʾ��ֻ���ڶᱦ����ʱ�ſ���ʹ��[ǧ���˲�]��[������֥]����ҩƷ��');}</script>"
		Response.End
	end if
end if
zswupin=trim(zt(0))
wusl=abs(int(clng(zt(1))))
if wusl=0 or wusl>100 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ����Ӧ����0С��100��');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select w1 from �û� where ����='"&sjjh_name&"'",conn,2,2
eattemp=abate(rs("w1"),zswupin,wusl)
rs.close
rs.open "select d,e FROM b WHERE a='" & zswupin &"'",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������Ʒ["&zswupin&"]��ϵͳ���ݿ��в�������\n��ɾ������Ʒ���ҹ���Ա����');}</script>"
	Response.End
end if
nl=abs(rs("d"))*wusl
tl=abs(rs("e"))*wusl
conn.execute "update �û� set w1='"&eattemp&"',����=����+"&nl&",����=����+"&tl&" where  ����='"&sjjh_name&"'"
eat="##ʹ����ҩƷ[<b>"&zswupin&"</b>]���ƣ�"&wusl&"��,�Լ���������<font color=red>+"&tl&"</font>��,����:<font color=red>+"&nl&"</font>��!"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>