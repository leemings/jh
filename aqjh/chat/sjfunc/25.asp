<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'��ʦ
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
if aqjh_jhdj<jhdj_bs then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ʦ��Ҫ["&jhdj_bs&"]���ſ��Բ�����');}</script>"
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
says="<font color=green>����ʦ��<font color=" & saycolor & ">"+bais(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��ʦ
function bais(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ʦ��,����,�ȼ� from �û� where ����='" & aqjh_name &"'",conn,2,2
sf=rs("ʦ��")
if aqjh_name=to1 then
	if sf="" or sf="��" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����û��ʦ�����޷�����ʦ�ţ�');}</script>"
		Response.End
	end if
if rs("����")<80000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û��8��飬�˼Ҳ���ѽ����');}</script>"
	Response.End
end if
	conn.execute "update �û� set ����=����-50000,ʦ��='��',����1='����' where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	bais="##��ԭʦ��"& sf &"������8���ķ��ַѣ�������"& sf &"������ʦͽ��ϵ��"
	exit function
end if
if sf=to1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1&"]�Ѿ�����ʦ����');}</script>"
	Response.End
end if
if rs("ʦ��")<>"" and rs("ʦ��")<>"��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����["& to1&"]Ϊʦ����������ʦ�������ϵ��');}</script>"
	Response.End
end if
if rs("����")<50000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����û��5����["& to1&"]��������Ϊ���ӣ�');}</script>"
	Response.End
end if
dj=rs("�ȼ�")
rs.close	
rs.open "select �ȼ� FROM �û� WHERE ����='" & to1 &"'",conn,2,2
if rs("�ȼ�")<30 or dj>rs("�ȼ�") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["& to1&"]������30��ѽ,�����յ��ӣ�');}</script>"
	Response.End
end if
bais="##��׼����%%������5����ʦ�ѣ�����˰�ʦ���룬Ҳ��֪���˼�ͬ�ⲻ��"
Application.Lock
Application("aqjh_bais_sf")=to1
Application("aqjh_bais_td")=aqjh_name
Application.UnLock
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
