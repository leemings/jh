<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
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
if aqjh_grade<grade_db then
	Response.Write "<script language=JavaScript>{alert('��ʾ����������ȼ�["&grade_db&"]�ſ��Բ�����');}</script>"
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
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
says=replace(says,chr(39),"")
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����¾�����</font><font color=" & saycolor & ">"+daipu(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'����
function daipu(fn1,to1)
if to1=Application("aqjh_user") then
	Response.Write "<script language=JavaScript>{alert('���棺�㲻���Զ�վ��������');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,grade,���� from �û� where ����='" & to1 &"'",conn
if aqjh_grade<=rs("grade") and aqjh_name<>Application("aqjh_user") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��ʧ�ܣ��㲻�ܶԸ߼��������!');}</script>"
	Response.End
end if
conn.execute "update �û� set ״̬='��',��¼=now()+3,�¼�ԭ��='"&aqjh_name& " ������" &fn1&"' where ����='" & to1 &"'"
conn.execute "insert into l(b,c,d,a,e) values ('" & aqjh_name & "','" & to1 & "','����',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
daipu="##:<font color=red><bgsound src=wav/oh_no.wav loop=1>�ٸ��������˸���������%%�Ĳ�����,���������İ�%%Ѻ�����η�3��ʱ�䣡" & fn1 & "</font>"
call boot(to1,"�����������ߣ�"&aqjh_name&","&fn1)
end function
%>