<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.asp"-->
<!--#include file="sjfunc.asp"-->
<%'��������
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

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
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
says=Replace(says,"\","�Ҹ���")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=#ff5600>���������ޡ�<font color=" & saycolor & ">"+tiren(towho,mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'��������
function tiren(to1,fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
myzs=rs("ת��")
zaohuan=rs("�ٻ���1")

if myzs<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�������������ô??');}</script>"
	Response.End
end if
if rs("ְҵ")<>"�ٻ�ʦ" then
	Response.Write "<script language=JavaScript>{alert('û����������л����ޣ�����ȥְҵת��Ϊ�ٻ�ʦ��');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if zaohuan="��" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�������������ô??');}</script>"
	Response.End
end if

if to1="���" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Ҫ����,��ķ���������֧����');}</script>"
	Response.End
end if
if to1=rs("����") then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��϶���ë��,���Լ�������');}</script>"
	Response.End
end if
if zaohuan<>to1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��Ҫ����,��ķ����������ٻ�����');}</script>"
	Response.End
end if
	tiren="<bgsound src=wav/py.wav loop=1>##����%%���εĸ����Լ����,���з���,��������,ֻ������%%����һ����Х,ƾ����ʧ�ڿ�����,ԭ���ǻص���������˯��ȥ��"
	call boot(to1,"��������&�������ߣ�"&aqjh_name&","&fn1)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>