<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../config.asp"-->
<%'����
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
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red>������<font color=" & saycolor & ">"+dazhuo()+"</font>"
towhoway=0
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,������,����,����ʱ�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
cy=DateDiff("n",rs("����ʱ��"),now())
if cy<7 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=7-cy
	Response.Write "<script language=JavaScript>{alert('�������["& ss &"]��,�ٲ�����');}</script>"
	Response.End
end if
if rs("����")<20000  then
	Response.Write "<script language=JavaScript>{alert('��20000���ϵ������ſ��ԣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("����")<rs("�ȼ�")*aqjh_nlsx+2000+rs("������") then
	conn.execute "update �û� set ֪��=֪��+50,����=����-10000,����ʱ��=now() where ����='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>�ڵ��ɽ����������<bgsound src=wav/dz.wav loop=1>,����ʧȥ-10000��֪������+50,�еñ���ʧ��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("������")<=rs("�ȼ�")*15000 then
	conn.execute "update �û� set ֪��=֪��+80,����=����-10000,����ʱ��=now() where ����='" & aqjh_name &"'"
	dazhuo="##<bgsound src=wav/dz.wav loop=1>�ڵ��ɽ����������<bgsound src=wav/dz.wav loop=1>,����ʧȥ<font color=red>-10000</font>��,֪����������<font color=red>+80</font>��,�еñ���ʧ��!"
else
	Response.Write "<script language=JavaScript>{alert('��������������ˣ������˼�������');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>