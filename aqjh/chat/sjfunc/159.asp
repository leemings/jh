<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'ħ������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻����ʹ��ħ�����');}</script>"
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
'call dianzan(towho)
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
says="<font color=green>��ħ��ʦ��ħ�����</font><font color=" & saycolor & ">"+zhouyu(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'ħ������
function zhouyu(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'",conn,3,3
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('���ǳ����˰���!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('������ɱ����������ʲô!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
f=Minute(time())
if f<00 or f>25 then
	Response.Write "<script language=JavaScript>{alert('�����ж�������Ķ���ÿСʱ��ǰ25���ӣ�����������ʱ�䣡');window.close();}</script>"
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
if rs("�ȼ�")<60 then
Response.Write "<script language=JavaScript>{alert('�˹�����Ҫ60��ս���ȼ�ѽ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")<50000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��50000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ���� FROM �û� WHERE ����='" & to1 &"'",conn
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԸ߼�����Ա��վ���ط����Ա����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ�����Գ�������ʲô��������!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �û� set ����=����-3000 where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����-8000 where ����='" & to1 &"'"
zhouyu=aqjh_name & "���÷��������<font color=red>" & to1 & "</font>��������ʧ,����������..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>t ����=����-8000 where ����='" & to1 &"'"
zhouyu=aqjh_name & "���÷��������<font color=red>" & to1 & "</font>��������ʧ,����������..." 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>