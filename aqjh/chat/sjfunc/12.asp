<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���Ǵ�
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_xx then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����Ǵ���Ҫ["&jhdj_xx&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>�����Ǵ󷨡�<font color=" & saycolor & ">"+xxdf(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���Ǵ�
function xxdf(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���� from �û� where ����='" & aqjh_name &"'" ,conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������˲����Բ���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ��ȡ����!');}</script>"
	Response.End
end if
rs.close
rs.open "select ����,����,�ȼ�,����,����,grade from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��ǳ����˲����Բ���!');}</script>"
	Response.End
end if
if rs("����")=True then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if
if rs("�ȼ�")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�벻Ҫ�Գ��뽭�����ֲ�����');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ҫ͵����Ա��Ǯ����ѽ������');}</script>"
	Response.End
end if
to1=rs("����")
nlto=rs("����")
rs.close
randomize timer
s=int(rnd*20)+1
nlto=int((nlto/100))*s
r=int(rnd*4)
Select Case r
Case 1
	xxdf="##<bgsound src=wav/xixing.wav loop=1>ʹ��һ�����Ǵ�,��ȡ%%����<font color=red>"& s &"%</font>,����:<font color=red>"& nlto &"</font>��!����С��ȫ����ʧ��!"
Case 2
	xxdf="##<bgsound src=wav/xixing.wav loop=1>����ȡ%%����,û�гɹ����Լ�ʧȥ����<font color=red>"& s &"%</font>,����:<font color=red>"& nlto &"</font>��!%%������Ц,֪���������˰ɣ�"
	conn.execute "update �û� set ����=����-"&nlto&" where ����='" & aqjh_name &"'"
case 3
	rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
	if rs("����")>50000 then
		conn.execute "update �û� set ����=����-50000 where ����='" & aqjh_name &"'"
		xxdf="##<bgsound src=wav/qt.wav loop=1>͵%%������������,ֻ�����´�㻨��<font color=red>5��</font>��,�ӹ��˽�,����ѽ!"
	else
		conn.execute "update �û� set ״̬='��' where ����='" & aqjh_name &"'"
		xxdf="##<bgsound src=wav/oh_no.wav loop=1>�ոհ��ִ���%%�����,�ͱ��ٸ����˷�����,������׽ȥ������"
		call boot(aqjh_name,"��ȡ�����������ߣ�"&to1&",����������η���")
	end if
	rs.close
case else
	xxdf="##<bgsound src=wav/xixing.wav loop=1>ʹ��һ�����Ǵ�,��ȡ%%����<font color=red>"& s &"%</font>,����:"& nlto &"��!����%%��Ѫ��"
	conn.execute "update �û� set ����=����+"&nlto&" where ����='" & aqjh_name &"'"
	conn.execute "update �û� set ����=����-"&nlto&" where ����='" & to1 &"'"
End Select
set rs=nothing	
conn.close
set conn=nothing
end function
%>
