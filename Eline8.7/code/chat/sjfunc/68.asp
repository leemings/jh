<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���ip��wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����ip��</font><font color=" & saycolor & ">"+jianjin(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'���ip
function jianjin(fn1,to1)
if sjjh_grade<10 or sjjh_name<>"���׵���" then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����IPֻ����վ���ſ��Բ�����');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
if to1=Application("sjjh_user") then
	call boot(sjjh_name,"���ip�������ߣ�"&sjjh_name&","&fn1)
	conn.execute "update �û� set ״̬='���',��¼=now()+1,�¼�ԭ��='���վ��ip��������' where ����='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
rs.open "select ����,grade,lastip from �û� where ����='" & to1 &"'",conn,2,2
if sjjh_grade<=rs("grade") and sjjh_name<>Application("sjjh_user")  then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻����վ���������Բ�������վ����');}</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
tolastip=rs("lastip")
conn.execute "update �û� set ״̬='���',��¼=now()+1,�¼�ԭ��='"&sjjh_name&" ���ip:"&to1&"���ܽ�����ӭ\n�����������ip��ͬ��' where grade<>10 and lastip='" & tolastip &"'"
jianjin= "<font color=red><bgsound src=wav/daipu.wav loop=1>��Ϊ%%�ڽ����ϲ��ܻ�ӭ��վ������������ͬ��¼ip�������ˡ���</font>"
call boot(to1,"����������ߣ�"&sjjh_name&","&fn1)
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','���',now(),'��" & to1 & "��ͬip�����˼��!')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>