<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'�ⶾ��
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
says="<font color=red>���ⶾ����<font color=" & saycolor & ">"+jds(mid(says,i+1),towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function jds(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,grade,����,����,����,ְҵ FROM �û� WHERE ����='" & aqjh_name &"'" ,conn,2,2
fn1=abs(fn1)
if fn1<100 then
Response.Write "<script language=JavaScript>{alert('�õĹ���̫�ͣ�');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
if to1="���" then
 jds=from1 & "��Ա��˳�����?!��������ʹ�ýⶾ��ѽ��"
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
	Response.Write "<script language=JavaScript>{alert('��ĵȼ�����ѽ��ʹ�ýⶾ��������Ҫ80����');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if fn1>200000 and rs("grade")<10 then
Response.Write "<script language=JavaScript>{alert('��Ҫ̫���������20��');}</script>"
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
if rs("����")<30000 then
Response.Write "<script language=JavaScript>{alert('�㷨������������Ҫ30000��');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,����,״̬ FROM �û� WHERE ����='"&to1&"'",conn
if rs("״̬")="����" then
jds=to1&"���˿���" & aqjh_name & "������ָ���ֿ��˿��Լ�����û�ж�ѽ����Ҫ����!"
else
conn.execute "update �û� set ����=����-100,����=����-" & fn1 & " where ����='" & aqjh_name &"'"
conn.execute "update �û� set ״̬='����' where ����='"&to1&"'"
jds=aqjh_name & "ʹ����" & fn1 & "������,�����Ǳ����"& to1 &"�����ѾõĶ��"&to1&"�м���ֱҪ����������"
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function

%>