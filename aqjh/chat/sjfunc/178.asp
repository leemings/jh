<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'͵����
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
says="<font color=red>��͵������<font color=" & saycolor & ">"+tqs(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
function tqs(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,grade,����,����,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 or fn1>200 then
Response.Write "<script language=JavaScript>{alert('һ������͵100��಻�ܳ���200��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="���" then
 tqs=from1 & "��Ա��˳�����?!��������ʹ��͵����ѽ��"
 conn.close
 set rs=nothing
 set conn=nothing
 exit function
end if
if rs("ְҵ")<>"ð�ռ�" then
	Response.Write "<script language=JavaScript>{alert('���ð�ռң������ұ������ȥְҵת��Ϊð�ռң�');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�ȼ�")<80 then
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ��͵����������Ҫ80����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>200 and rs("grade")<10 then
	Response.Write "<script language=JavaScript>{alert('͵�������ܴ���200��');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if rs("����")<fn1 then
Response.Write "<script language=JavaScript>{alert('��������ô�������������ˣ�');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if rs("����")<25000 then
	Response.Write "<script language=JavaScript>{alert('�㷨������500�ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ���� FROM �û� WHERE ����='"&to1&"'",conn
if rs("����")<fn1 then
tqs=to1&"�������еĿ޵���" & aqjh_name & "���кðɣ���������ô�������ѽ�����˰���,����!"
else
conn.execute "update �û� set ����=����-500,����=����-"& fn1/1 & ",����=����+" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ����=����-" & fn1 & " where ����='"&to1&"'"
tqs="����"& to1 & "�����,һ���ѱ�ð׳հ׳յ�ԭ������������" & fn1 & "���������ѹֿ������е�׳ա���(" & aqjh_name & "�����½�"& fn1/1 &")"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>