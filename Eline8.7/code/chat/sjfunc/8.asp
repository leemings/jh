<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'ն�ס�wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
useronlinename=Application("sjjh_useronlinename"&nowinroom)
if Instr(LCase(useronlinename),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if sjjh_grade<grade_zs then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ����ȼ�["&grade_zs&"]�ſ��Բ�����');}</script>"
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
says="<font color=green>��ն�ס�<font color=" & saycolor & ">"+zhanshou(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'ն��
function zhanshou(fn1,to1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("sjjh_usermdb")
rs.open "select ����,grade from �û� where ����='" & to1 &"'",conn,2,2
if sjjh_grade<rs("grade") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�㲻���ԶԸ߼�����Ա������');}</script>"
	Response.End
end if
conn.execute "update �û� set ״̬='��',allvalue=0,grade=1,����='����',��Ա�ȼ�=0,��Ա��=0,mvalue=0,�¼�ԭ��='"&sjjh_name&" ն��"&fn1&"' where ����='" & to1 &"'"
zhanshou= "<font color=red>�ٸ����˺���:�˷�%%ն����!</font>"& fn1
call boot(to1,"ն�����������ߣ�"&sjjh_name&","&fn1)
conn.execute "insert into l(b,c,d,a,e) values ('" & sjjh_name & "','" & to1 & "','ն��',now(),'" & fn1 & "')"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>