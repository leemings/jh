<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'Ǭ��һ����wWw.51eline.com��
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

'�����뿪������Ѩ�ж�
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
says="<font color=green>��Ǭ��һ����<font color=" & saycolor & ">"+moneygj(mid(says,i+1)+0,towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'Ǭ��һ��
function moneygj(fn1,to1)
fn1=int(abs(fn1))
if (fn1<10000 or fn1>50000000) and sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��Ǭ��һ������1���������5000������');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����ʱ��,grade,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=2-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ���������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if

if rs("����")<fn1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ô���Ǯ�𣿿�����˵��');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ���Ա!');}</script>"
	Response.End
end if
if rs("����")=True and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ڱ�����!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,grade,����,ͨ�� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("����")="����" and sjjh_grade<10 and rs("ͨ��")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɹ���������!');}</script>"
	Response.End
end if
if rs("grade")>=6 and sjjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ɹ����ٸ���Ա!');}</script>"
	Response.End
end if
if rs("����")=True and sjjh_grade<10 and rs("ͨ��")=False then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է����ڱ�����!');}</script>"
	Response.End
end if
if rs("����")<=1000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&to1&"]�������������ˣ��㲻Ҫ�ٹ�����!');}</script>"
	Response.End
end if
randomize ()
m1 = Int(100 * Rnd)+100
gjtl=int(fn1/m1)
gjtl=rs("����")-gjtl
if gjtl<1000 then 	gjtl=1000
conn.execute "update �û� set ����="&gjtl&" where ����='" & to1 &"'"
conn.execute "update �û� set ����=����-" & fn1 & ",����ʱ��=now() where ����='" & sjjh_name &"'"
moneygj="##�ӿڴ����ó�����" & fn1 & "����ʹ����һ���������Ǭ��һ��,���е�������һ�����ε����������%%��ͷת��%%����ֻʣ��[$$redb"&gjtl&"$$b]��~~"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
