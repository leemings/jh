<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��ʩ����wWw.happyjh.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Բ�ʩ����');}</script>"
	Response.End
end if
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
'call dianzan(towho)
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if sjjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("sjjh_automanname") or towho=sjjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>����ʩ���¡�</font><font color=" & saycolor & ">"+bushishu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ʩ��
function bushishu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ����,����,�ȼ� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("����")<5000000 then
Response.Write "<script language=JavaScript>{alert('������û��500���ȴ�500����������˵�ɣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<5000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��5000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�ȼ�")<25 then
Response.Write "<script language=JavaScript>{alert('��˹�����Ҫ25��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-5000,����=����+1000,����=����-5000000 where ����='"&sjjh_name&"'"
conn.execute "update �û� set ����=����+2000000 where ����='" & to1 &"'"
rs.close
rs.open "select �ȼ�,���� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
ddj=(rs("�ȼ�")*1440+2200)+rs("����")
if rs("����")>=ddj then
conn.execute "update �û� set ����="&ddj&" where ����='"&sjjh_name&"'"
end if
bushishu=sjjh_name & "���÷���һʽ��ʩ���·���<font color=red>" & to1 & "</font>����500�������Լ���������1000��." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>