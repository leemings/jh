<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'������Ʒ
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
says=jbsp(mid(says,i+1))
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'������Ʒ
function jbsp(fn1)
fn1=trim(fn1)
zt=split(fn1,"|")
if ubound(zt)<>1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʽΪ��������Ʒ��|��Ǯ ');}</script>"
	Response.End 
end if
'ȡ��������Ʒ��
wpname=trim(zt(0))
money=abs(int(clng(zt(1))))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if Weekday(date())<>1 or (Hour(time())<20 or Hour(time())>=21) then
'if Weekday(date())=7 then
	rs.open "select r FROM b WHERE a='����'",conn,2,2
	'�ж��Ƿ���ɾ��꣬�����ȷ�����ꡣ
	if rs("r")=True then
		conn.execute "update b set r=False where a='����'"
		rs.close
		rs.open "select a,n,q,o FROM b WHERE n<>'' and q>=10000 and r=True",conn,2,2	
		do while not rs.bof and not rs.eof
			jbname=jbname&"<font color=blue>"&rs("n")&"</font>��Ӫ[<font color=blue><b>"&rs("a")&"</b></font>]Ͷ�ʼ�:<font color=red>"&rs("q")&"��</font>  "
			conn.execute "update �û� set ���=���-"&rs("q")&" where ����='"&rs("n")&"'"
			rs.movenext
		loop
		'������о�����ƷΪ�Ѿ���(�����δ������򲻱��)
		conn.execute "update b set r=False,p='��Ӫ��'&n&'˵:��ӭ���ѡ���ҵ���Ʒ��лл֧��!' where n<>'' and q>=10000"
		conn.execute "update b set r=False"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		jbsp="<font color=green><b>��������Ʒ�����</b><font color=" & saycolor & ">"&jbname&"��Ʒ������ɣ���λ��þ�ӪȨ����ҿ��Թ����ˡ���</font>"
		exit function
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ʒֻ����������20-21����У�');}</script>"
	Response.End
end if
'���Ѿ��������Ʒ����
rs.open "select r FROM b WHERE a='����'",conn,2,2
'�ж��Ƿ���ɾ��꣬�����ȷ�����ꡣ
if rs("r")=False then
	conn.execute "update b set r=True where a='����'"
	rs.close
	rs.open "select a,n,q,o FROM b WHERE n<>'' and o>=1 and r=False",conn,2,2	
	do while not rs.bof and not rs.eof
		jbname=jbname&"��Ӫ[<font color=blue><b>"&rs("a")&"</b></font>]��<font color=blue>"&rs("n")&"</font>����:<font color=red>"&rs("o")&"��</font>  "
		conn.execute "update �û� set ���=���+"&rs("o")&" where ����='"&rs("n")&"'"
		rs.movenext
	loop
	'������о�����ƷΪ�Ѿ���(�����δ������򲻱��)
	conn.execute "update b set r=True,n='��',q=0,o=0,p='������!' where b='�ʻ�' or b='ҩƷ'"
	conn.execute "update b set r=True"
	rs.close
	set rs=nothing
	conn.close 
	set conn=nothing
	jbsp="<font color=green><b>��������Ʒ�ֺ졿</b><font color=" & saycolor & ">"&jbname&"</font>���ڽ�����λ��ҿ��Ծ����Լ�ϲ������Ʒ�ˡ�"
	exit function
end if
'���о���ϵͳ
rs.close
rs.open "select a,n,q,o FROM b WHERE n='"&aqjh_name&"'",conn,2,2	
trmoney=0
do while not rs.bof and not rs.eof
	trmoney=trmoney+rs("q")
	rs.movenext
loop
rs.close
rs.open "select ��� FROM �û� WHERE ����='"&aqjh_name&"'",conn,2,2
if rs("���")<(money+trmoney) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	if trmoney=0 then
		Response.Write "<script language=JavaScript>{alert('��ʾ����Ĵ���["&money&"]��');}</script>"
	else
		Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ���Ͷ�������ڵ�Ǯ�������Ѿ�Ͷ��ģ�����["&money&"]��');}</script>"
	end if
	Response.End
end if
rs.close
rs.open "select a,n,q FROM b WHERE a='"&wpname&"' and (b='ҩƷ' or b='�ʻ�') ",conn,2,2
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ʒ��������ݳ���!');}</script>"
	Response.End
end if
if aqjh_name=rs("n") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������������㾺��ģ�!');}</script>"
	Response.End
end if
if money<rs("q") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������۸񲻹�!');}</script>"
	Response.End
end if
conn.execute "update b set n='"&aqjh_name&"',q="&money&",o=0 where a='"&wpname&"'"
jbsp="<font color=green>��������Ʒ��<font color=" & saycolor & ">##������Ʒ[<b>"&wpname&"</b>]�ľ�ӪȨ���˴γ���:+<font color=red>"&money&"</font>ϣ�����Ի�þ�ӪȨ����</font>"
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
