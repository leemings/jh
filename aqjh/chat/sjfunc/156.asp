<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%'����׶
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
says="<font color=green>������׶��<font color=" & sayscolor & ">"+potianzhui(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function potianzhui(fn1)
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
if rs("����")<1000 then
Response.Write "<script language=JavaScript>{alert('��ķ��������޷�ʩչѽ������Ҳ��1000�㰡��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ���� FROM �û� WHERE  ����="&"'" & towho &"'",conn,3,3
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
wuyn=iswp(w6w,"����׶")
if wuyn=0 then
Response.Write "<script language=JavaScript>{alert('��������׶��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
fq=abate(w6w,"����׶",1)
conn.execute "update �û� set  w9='"&fq&"' where ����='"&aqjh_name&"'"
conn.execute "update �û� set ����=����-1000 where  ����='" & aqjh_name & "'"
potianzhui="##����ħ��<bgsound src=wav/Bombs020.wav loop=1>����׶����<font color=red>%%</font>����һ׶,��%%����������ң��εĺ���..." 
 call boot(towho,"����׶�������ߣ�"&aqjh_name&","&fn1)
rs.close
set rs=nothing	
	conn.close
	set conn=nothing	
end function 
%>