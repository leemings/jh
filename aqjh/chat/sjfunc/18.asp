<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'�鿴״̬
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
if aqjh_grade<grade_see then
	Response.Write "<script language=JavaScript>{alert('��ʾ���鿴״̬��Ҫ����ȼ�["&grade_see&"]�ſ��Բ�����');}</script>"
	Response.End
end if
call dianzan(towho)
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>1500 then says=Left(says,1500)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=1
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>��״̬��<font color=" & saycolor & ">##�õ������鱨��"+getstat(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'�鿴״̬
function getstat(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������鿴���˲����ڣ�����');}</script>"
	Response.End
end if
if to1=Application("aqjh_user") then
    	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������վ��,������������');}</script>"
	Response.End
end if
getstat=to1 &"��" & rs("�Ա�") & "�����£�" & rs("��ż") & "�����ˣ�" & rs("����") & "��ʦ����" & rs("ʦ��") & "�����ɣ�" &rs("����") & "����ݣ�" &rs("���") &  "���书��" &rs("�书") & "��������" & rs("����") & "��������" & rs("����") & "����������" & rs("����") & "����������" & rs("����")&"����¼ip��" & rs("lastip") & "�����䣺" & rs("����") & "��������" & rs("����") & "���Ĵ�����" & rs("�Ĵ���") & "��Ӯ������" & rs("Ӯ����") & "��ӮǮ��" & rs("ӮǮ") & "�����£�" & rs("����") & "������ȼ���" & rs("grade") & "���ۻ��֣�" & rs("allvalue") & "���»��֣�" & rs("mvalue") & "����¼��" & rs("��¼") & "��������" & rs("����")&"����ң�"&rs("���")&"����"&rs("���")&"�������ˣ�"&rs("������")
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>