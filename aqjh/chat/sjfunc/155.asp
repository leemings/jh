<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'Ѫ����
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
%>
<!--#include file="pkif.asp"-->
<%
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
says="<font color=green>��Ѫ���ӡ�<font color=" & sayscolor & ">"+xuedizi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function xuedizi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w6,w9,����,����,�ȼ� FROM �û� WHERE  ����="&"'" & aqjh_name &"'",conn,3,3
w6w=rs("w9")
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('���ǹٸ���Ա��!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('���ǳ����˲��ܲ�����!');}</script>"
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
if rs("����")<8000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��8000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ���� FROM �û� WHERE ����="&"'" & towho &"'",conn,3,3
if rs("����")="�ٸ�"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԸ߼�����Ա��վ���ط����Ա����!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("����")="����"  then
Response.Write "<script language=JavaScript>{alert('ʧ�ܣ��㲻�ܶԳ����˲���!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
wuyn=iswp(w6w,"Ѫ����")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('����Ѫ������');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"Ѫ����",1)
conn.execute "update �û� set  w9='"&fq&"' where ����='"&aqjh_name&"'"

conn.execute "update �û� set ����=����-8000 where  ����="&"'" & aqjh_name &"'"

conn.execute "update �û� set ����=����-2000,����=0 where ����="&"'" & towho &"'"
conn.execute "update �û� set ����=����-2000 where  ����="&"'" & towho &"'"
rs.open "select * from �û� where ����="&"'" & towho &"'",conn
if rs("����")<0  then
conn.execute "update �û� set ɱ����=ɱ����+1 where ����="&"'" & aqjh_name &"'"
conn.execute "update �û� set ״̬='��' where ����="&"'" & towho &"'"
xuedizi="##�׳�����Ѫ���ӣ�<bgsound src=wav/Bombs020.wav loop=1>Ѫ����������<font color=red>%%</font>ͷ�ϣ�%%��Ѫ���ӻ��з���ȫ�ޣ���������2000������ʧȥ2000����������..." 
 call boot(towho,"Ѫ���ӣ������ߣ�"&aqjh_name&","&fn1)
else
xuedizi="##�׳�ħ��<bgsound src=wav/Bombs020.wav loop=1>Ѫ���ӣ�Ѫ����������<font color=red>%%</font>ͷ�ϣ�%%��Ѫ���ӻ��з���ȫ�ޣ���������2000������ʧȥ2000��û���Ѿ���������..." 
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>