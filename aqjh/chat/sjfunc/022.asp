<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'���齵��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl ="No-Cache"
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
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=fakuan(mid(says,i+1)+0,towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'���齵��"
function fakuan(fn1,to1)
fn1=abs(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ʦ��,����,����,ɱ����,��ɱ��,sl,��Ա�� FROM �û� WHERE ����='" & aqjh_name &"'",conn,2,2
if rs("����")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>С���</font>]��[<font color=red>"&aqjh_name&"</font>]˵����������ô���Ǯ������ˣ�����һ��������"
	exit function
end if
if rs("sl")<>" " and rs("sl")<>"��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
      fakuan="<font color=red>�����齵����</font>������ʾ��[<font color=red>"&aqjh_name&"</font>]���Ѿ�����������ˣ���Ҫ�����ˣ��𿨶���͸����˰ɣ�"
	exit function
end if
if rs("��Ա��")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="[<font color=red>�����</font>]��[<font color=red>"&aqjh_name&"</font>]˵����������<img src='picwords/1.gif'><img src='picwords/0.gif'>��𿨣�����ˣ�����һ��������"
	exit function
end if
rs.close
rs.open "SELECT ����,grade,ʦ��,����,��ɱ��,sl,slsj,��Ա��,֪��,���� FROM �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("��ɱ��")>20 then
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    fakuan=""&aqjh_name&"��������[����]���ٸ�����������"&aqjh_name&"������������ɱ��������<img src='picwords/2.gif'><img src='picwords/0.gif'>�ˣ��벻��[����]���ٰ�æ��"
	exit function
end if
if rs("����")>=fn1 then
    conn.execute("update �û� set ֪��=֪��+10,����=����+10,sl='����',slsj=now()+2,��Ա��=��Ա��-10,����=����-" & fn1 & " where ����='"&aqjh_name&"'")
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	fakuan="<font color=" & saycolor & "><font color=red>�����齵����</font>"&aqjh_name&"��һ�����������һ��[<font color=red>�����</font>]������" & fn1 & "�������Ű��ɵ���<font color=red>10</font>��𿨼�����~[<b><font color=red>����</font></b>]��Ľ���"&aqjh_name&"���ϣ��ݵ��ٶ�����2����ʱ��Ϊ1�죬���鸽�����������10�㣬֪������10�㣡"
	exit function
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>
