<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'NPC�ٻ���
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
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
says="<font color=green>��NPC�ٻ���<font color=" & saycolor & ">"+zhnpc(mid(says,i+1),towho)+"</font>"
call chatsay("NPC�ٻ�",towhoway,towho,saycolor,addwordcolor,addsays,says)
function zhnpc(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'��ʼ�ж�npc
rs.open "select * from npc where n����='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��NPC["&to1&"]�������ڻ���������NPC�����˰ɣ�');}</script>"
	Response.End	
end if
if InStr(";" & Application("aqjh_npc"), ";" &to1& "|")=0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��NPC["&to1&"]�����ߣ�');}</script>"
	Response.End
end if
if rs("n����")=aqjh_name then
  zhnpc="##:����NPC[%%]�ĵ��ˣ����������һ���ߵģ����߿��ɣ�"
  exit function
end if
n_zhuren=rs("n����")
rs.close
'��ʼ�ж�npc������û����
if n_zhuren<>"��" then
rs.open "select * from �û� where ����='" &n_zhuren&"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������["&z_zhuren&"]������,NPC�����쳣��');}</script>"
	Response.End	
end if
if rs("״̬")<>"����" then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��NPC["&to1&"]������["&z_zhuren&"]��û�����أ�');}</script>"
	Response.End
end if
elseif n_zhuren="��" then
 	Response.Write "<script language=JavaScript>{alert('��ʾ���ٸ����ŵ�NPC���������ٻ���');}</script>"
	Response.End
end if
if n_zhuren=aqjh_name then
   zhnpc="��ʾ��##:���ʲô�������Ѿ���NPC[%%]�������ˣ�"
   exit function
end if
'�ж��Լ�
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")<10000 or rs("����")<10000 then
   zhnpc="##���%%ʹ��[�ٻ���]���������Ժ�....�����Լ�����̫ǳ�˴β���ʧ����,�װ�����ԭ��!"
   exit function
end if
if rs("���")<5 then
   zhnpc="��ʾ��##:���û��5����NPC�������ߣ�"
   exit function
end if
if rs("ְҵ")<>"ħ��ʦ" then
   zhnpc="��ʾ��##:�㲻��ħ��ʦ����û���ٻ��Ĺ��ܣ�"
   exit function
end if
conn.execute "update npc set n����='��',n����='"&aqjh_name&"' where n����='"&to1&"'"
conn.execute "update �û� set ����=����-10000,����=����-10000,���=���-5 where ����='"&aqjh_name&"'"
zhnpc="[##]��NPC����[%%]ʩչ[�ٻ���]���Թ����������һ�һ��û��ģ����ɣ�����......(NPC����[%%]ʧȥ����־��##����%%������)"
rs.close
conn.close
end function
%>