<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'��ʩ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Բ�ʩ����');}</script>"
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
act=0
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
if trim(towho)="" or towho="���" or towho=application("aqjh_automanname") or towho=aqjh_name then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻���ԶԴ�ҡ������˻����ѽ��д��������');}</script>"
	Response.End
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
if i=0 then i=len(says)+1
mycz=trim(left(says,i-1))
says="<font color=green>��ħ��ʦ����ʩ���¡�</font><font color=" & saycolor & ">"+bushishu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'��ʩ��
function bushishu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,ְҵ,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<3000000 then
Response.Write "<script language=JavaScript>{alert('������û��300���ȴ�300����������˵�ɣ�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("ְҵ")<>"ħ��ʦ" then
	Response.Write "<script language=JavaScript>{alert('��ֻ����ͨ�����أ�����ȥְҵת��Ϊħ��ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���"& s &"���ٲ���,�ɱ����ţ�');</script>"
	response.end
end if
if rs("����")<2000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��2000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-2000,����=����+3000,����=����-3000000,����ʱ��=now() where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����+2000000 where ����='" & to1 &"'"
rs.close
rs.open "select �ȼ�,���� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
ddj=(rs("�ȼ�")*1440+2200)+rs("����")
if rs("����")>=ddj then
conn.execute "update �û� set ����="&ddj&" where ����='"&aqjh_name&"'"
end if
bushishu=aqjh_name & "���÷���һʽ��ʩ���·���<font color=red>" & to1 & "</font>����200�������Լ���ʧ����300�����������½�1000�㣬�Լ���������3000��." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>f'><img src='picwords/0.gif'>�����������½�<img src='picwords/1.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>�㣬�Լ���������<img src='picwords/3.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'><img src='picwords/0.gif'>��." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>